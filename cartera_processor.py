import io
import re
import urllib.parse
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd
import requests

from config import (
    CLIENTES_ELITE_PATH,
    DIAS_CARTERA_ALTO_IMPACTO_MIN,
    DIAS_CARTERA_PREMATURO_MAX,
    GOOGLE_SHEETS_NO_CONTACTAR_URL,
    MENSAJE_DEBITO_FALLIDO_TEMPLATE,
    NO_CONTACTAR_CACHE_PATH,
    SALDO_ALTO_IMPACTO_MIN,
)


def clean_id_value(val) -> str:
    """Normaliza un valor de identificación a string limpio sin decimales, puntos, comas ni espacios."""
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ("nan", "none", "null", "<na>"):
        return ""
    # Si viene con formato float como "123456.0"
    s = re.sub(r"\.0$", "", s)
    # Limpiar puntos y comas de formato de miles (ej. 42.680.537 -> 42680537)
    s = re.sub(r"[\.,\s]", "", s)
    # Remover prefijos como CC, NIT, CE, TI, etc.
    s = re.sub(r"^(cc|nit|ce|ti|ppt|pas)\s*", "", s, flags=re.IGNORECASE)
    # Si viene con dígito de verificación al final ej: 900123456-1
    if "-" in s:
        s = s.split("-")[0]
    return s.strip()


def extract_first_name(full_name) -> str:
    """Extrae el primer nombre propio de una persona en formato título legible."""
    if pd.isna(full_name) or full_name is None:
        return "estimado cliente"
    s = str(full_name).strip()
    if not s or s.lower() in ("nan", "none", "null", "<na>"):
        return "estimado cliente"
    parts = s.split()
    if not parts:
        return "estimado cliente"

    # Si empieza por títulos o tratamientos, tomar la siguiente palabra
    honorifics = {"dr", "dr.", "dra", "dra.", "sr", "sr.", "sra", "sra.", "ing", "ing.", "abg", "abg."}
    first = parts[1].strip() if len(parts) > 1 and parts[0].lower() in honorifics else parts[0].strip()

    return first.capitalize()


def clean_policy_value(val) -> str:
    """Normaliza un número de póliza a string limpio conservando ceros a la izquierda."""
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.lower() in ("nan", "none", "null", "<na>"):
        return ""
    s = re.sub(r"\.0$", "", s)
    return s.strip()


def parse_numeric(series: pd.Series) -> pd.Series:
    """Convierte una serie a numérico flotante, tolerando signos monetarios y separadores."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0.0)

    def _convert(val):
        if pd.isna(val) or val is None:
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val).strip()
        # Eliminar signos de moneda y espacios
        s = re.sub(r"[\$\sCOPcop]", "", s)
        if not s:
            return 0.0
        # Manejo de separadores miles y decimales
        if "." in s and "," in s:
            # Ejemplo: 1.234.567,89 (estándar colombiano/europeo)
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                # Ejemplo: 1,234,567.89 (estándar anglosajón)
                s = s.replace(",", "")
        elif "," in s:
            # Si solo tiene coma, determinar si es decimal o miles
            parts = s.split(",")
            if len(parts) == 2 and len(parts[1]) in (1, 2):
                s = s.replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "." in s:
            # Si solo tiene puntos, verificar si es miles ej: 3.000.000
            parts = s.split(".")
            if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3):
                s = s.replace(".", "")

        try:
            return float(s)
        except ValueError:
            return 0.0

    return series.apply(_convert)


def find_column_name(df: pd.DataFrame, target_name: str, alternatives: Optional[List[str]] = None) -> str:
    """Busca una columna de forma flexible tolerando tildes, mayúsculas y espacios."""
    if target_name in df.columns:
        return target_name

    all_candidates = [target_name] + (alternatives or [])
    norm_map = {re.sub(r"[\s\._º°]", "", col).lower(): col for col in df.columns}

    for cand in all_candidates:
        clean_cand = re.sub(r"[\s\._º°]", "", cand).lower()
        if clean_cand in norm_map:
            return norm_map[clean_cand]

    # Búsqueda por subcadena si no se halló coincidencia exacta
    for cand in all_candidates:
        clean_cand = re.sub(r"[\s\._º°]", "", cand).lower()
        for norm_col, orig_col in norm_map.items():
            if clean_cand in norm_col or norm_col in clean_cand:
                return orig_col

    return target_name


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    """Lee un archivo cargado en memoria (.xlsx o .csv) detectando formato por extensión o firma binaria."""
    file_name = getattr(uploaded_file, "name", "").lower()
    content_bytes = uploaded_file.getvalue() if hasattr(uploaded_file, "getvalue") else uploaded_file.read()

    # Detectar si es un archivo Excel (.xlsx / .xls) por extensión o por firma mágica (PK.. para zip/xlsx o OLE para xls)
    is_xlsx = file_name.endswith(".xlsx") or content_bytes[:4] == b"PK\x03\x04"
    is_xls = file_name.endswith(".xls") or content_bytes[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

    if is_xlsx or is_xls:
        return pd.read_excel(io.BytesIO(content_bytes), dtype=str)

    # Es CSV o texto plano, detectar codificación y separador
    encodings = ["utf-8", "latin1", "cp1252", "iso-8859-1"]
    separators = [",", ";", "\t", "|"]

    for enc in encodings:
        try:
            decoded = content_bytes.decode(enc)
            # Detectar separador común en la primera línea
            first_line = decoded.splitlines()[0] if decoded.splitlines() else ""
            best_sep = max(separators, key=lambda s: first_line.count(s))
            if first_line.count(best_sep) > 0:
                return pd.read_csv(io.StringIO(decoded), sep=best_sep, dtype=str)
            return pd.read_csv(io.StringIO(decoded), dtype=str)
        except (UnicodeDecodeError, pd.errors.ParserError):
            continue

    # Fallback genérico
    return pd.read_csv(io.BytesIO(content_bytes), dtype=str, on_bad_lines="skip")


class ClientesEliteManager:
    """Gestiona la persistencia y lectura de clientes élite."""

    def __init__(self, file_path: Path = CLIENTES_ELITE_PATH):
        self.file_path = Path(file_path)
        self.ensure_default_file()

    def ensure_default_file(self):
        """Asegura que el archivo y su directorio existan."""
        self.file_path.parent.mkdir(parents=True, exist_ok=True)
        if not self.file_path.exists():
            df_default = pd.DataFrame({"Nº ident.fis.1": []})
            df_default.to_csv(self.file_path, index=False, encoding="utf-8")

    def load_elite_set(self) -> Set[str]:
        """Carga el conjunto de documentos élite como strings limpios."""
        self.ensure_default_file()
        try:
            df = pd.read_csv(self.file_path, dtype=str)
            col_id = find_column_name(df, "Nº ident.fis.1", ["Documento", "NIT", "Identificacion", "Cedula"])
            if col_id in df.columns:
                return {clean_id_value(v) for v in df[col_id] if clean_id_value(v)}
            # Si no encuentra el nombre exacto, tomar la primera columna
            if not df.empty:
                return {clean_id_value(v) for v in df.iloc[:, 0] if clean_id_value(v)}
            return set()
        except Exception:
            return set()

    def count(self) -> int:
        """Devuelve el total de clientes élite únicos."""
        return len(self.load_elite_set())

    def update_from_file(self, uploaded_file) -> Tuple[bool, str, int]:
        """Actualiza permanentemente el listado de clientes élite desde un archivo cargado."""
        try:
            df = read_uploaded_file(uploaded_file)
            col_id = find_column_name(df, "Nº ident.fis.1", ["Documento", "NIT", "Identificacion", "Cedula", "Numero"])
            if col_id in df.columns:
                series_id = df[col_id]
            else:
                series_id = df.iloc[:, 0]

            cleaned_ids = [clean_id_value(x) for x in series_id if clean_id_value(x)]
            unique_ids = sorted(list(set(cleaned_ids)))

            if not unique_ids:
                return False, "No se encontraron identificaciones válidas en el archivo.", 0

            df_to_save = pd.DataFrame({"Nº ident.fis.1": unique_ids})
            self.file_path.parent.mkdir(parents=True, exist_ok=True)
            df_to_save.to_csv(self.file_path, index=False, encoding="utf-8")

            return True, f"Se actualizaron exitosamente {len(unique_ids)} clientes élite.", len(unique_ids)
        except Exception as e:
            return False, f"Error al procesar el archivo de clientes élite: {str(e)}", 0


class NoContactarService:
    """Gestiona la obtención de clientes en lista de No Contactar desde Google Sheets o fallback local."""

    def __init__(self, url: str = GOOGLE_SHEETS_NO_CONTACTAR_URL, cache_path: Path = NO_CONTACTAR_CACHE_PATH):
        self.url = url
        self.cache_path = Path(cache_path)
        self.ensure_default_file()

    def ensure_default_file(self):
        """Asegura que el archivo de caché local exista."""
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        if not self.cache_path.exists():
            df_default = pd.DataFrame({"Nº ident.fis.1": []})
            df_default.to_csv(self.cache_path, index=False, encoding="utf-8")

    def count(self) -> int:
        """Devuelve el total de clientes a no contactar registrados en la caché local."""
        self.ensure_default_file()
        try:
            df = pd.read_csv(self.cache_path, dtype=str)
            col = find_column_name(df, "Nº ident.fis.1", ["Documento", "NIT", "Identificacion", "Cedula"])
            series = df[col] if col in df.columns else df.iloc[:, 0]
            ids = {clean_id_value(v) for v in series if clean_id_value(v)}
            return len(ids)
        except Exception:
            return 0

    def update_from_file(self, uploaded_file) -> Tuple[bool, str, int]:
        """Actualiza permanentemente el listado de clientes a no contactar desde un archivo."""
        try:
            df = read_uploaded_file(uploaded_file)
            col_id = find_column_name(df, "Nº ident.fis.1", ["Documento", "NIT", "Identificacion", "Cedula", "Numero"])
            series_id = df[col_id] if col_id in df.columns else df.iloc[:, 0]

            cleaned_ids = [clean_id_value(x) for x in series_id if clean_id_value(x)]
            unique_ids = sorted(list(set(cleaned_ids)))

            if not unique_ids:
                return False, "No se encontraron identificaciones válidas en el archivo.", 0

            df_to_save = pd.DataFrame({"Nº ident.fis.1": unique_ids})
            self.cache_path.parent.mkdir(parents=True, exist_ok=True)
            df_to_save.to_csv(self.cache_path, index=False, encoding="utf-8")

            return True, f"Se actualizaron exitosamente {len(unique_ids)} clientes en la lista de No Contactar.", len(unique_ids)
        except Exception as e:
            return False, f"Error al procesar el archivo de no contactar: {str(e)}", 0

    def fetch_no_contactar(self, manual_file=None) -> Tuple[Set[str], str, str]:
        """
        Obtiene la lista de identificaciones a no contactar.
        Retorna: (set_de_ids, fuente_utilizada, estado_o_advertencia)
        """
        # 1. Si el usuario suministró un archivo manual de reemplazo en la sesión
        if manual_file is not None:
            try:
                df_manual = read_uploaded_file(manual_file)
                col = find_column_name(df_manual, "Nº ident.fis.1", ["Documento", "NIT", "Identificacion"])
                series = df_manual[col] if col in df_manual.columns else df_manual.iloc[:, 0]
                ids = {clean_id_value(v) for v in series if clean_id_value(v)}
                # Actualizar caché permanente
                self.cache_path.parent.mkdir(parents=True, exist_ok=True)
                pd.DataFrame({"Nº ident.fis.1": list(ids)}).to_csv(self.cache_path, index=False, encoding="utf-8")
                return ids, "Archivo cargado en sesión", "OK"
            except Exception:
                pass

        # 2. Intentar consultar Google Sheets en línea (usando verify=False para tolerar certificados locales de Windows)
        try:
            import urllib3
            urllib3.disable_warnings()

            resp = requests.get(self.url, timeout=7, headers={"User-Agent": "GuiaBienCarteraApp/1.0"}, verify=False)
            content_type = resp.headers.get("Content-Type", "").lower()
            is_html = "text/html" in content_type or "<!doctype html" in resp.text[:200].lower()

            if resp.status_code == 200 and not is_html:
                df_sheet = pd.read_csv(io.StringIO(resp.text), dtype=str)
                series = df_sheet.iloc[:, 0]
                ids = {clean_id_value(v) for v in series if clean_id_value(v)}
                # Actualizar caché permanente
                self.cache_path.parent.mkdir(parents=True, exist_ok=True)
                pd.DataFrame({"Nº ident.fis.1": list(ids)}).to_csv(self.cache_path, index=False, encoding="utf-8")
                return ids, "Google Sheets (Sincronización en vivo)", "OK"
            else:
                warning_msg = (
                    "El enlace de Google Sheets es privado (requiere login o devolvió código "
                    f"{resp.status_code})."
                )
        except Exception as e:
            warning_msg = f"No se pudo conectar a Google Sheets ({type(e).__name__})."

        # 3. Fallback a caché local persistente
        if self.cache_path.exists():
            try:
                df_cache = pd.read_csv(self.cache_path, dtype=str)
                series = df_cache.iloc[:, 0]
                ids = {clean_id_value(v) for v in series if clean_id_value(v)}
                if ids:
                    return (
                        ids,
                        "Caché local persistente (Respaldo)",
                        f"{warning_msg} Se utilizó la lista almacenada ({len(ids)} clientes).",
                    )
            except Exception:
                pass

        return set(), "Sin lista", f"{warning_msg} No se aplicaron exclusiones de no contactar."


class CarteraPipeline:
    """Ejecutor del flujo de depuración y lógica de negocio de cartera."""

    def __init__(self, elite_manager: Optional[ClientesEliteManager] = None, no_contactar_service: Optional[NoContactarService] = None):
        self.elite_manager = elite_manager or ClientesEliteManager()
        self.no_contactar_service = no_contactar_service or NoContactarService()

    def process(
        self,
        df_cartera_actual_raw: pd.DataFrame,
        df_cartera_anterior_raw: pd.DataFrame,
        df_recaudos_raw: pd.DataFrame,
        manual_no_contactar_file=None,
    ) -> Dict[str, any]:
        """
        Ejecuta el pipeline completo de depuración de cartera según las especificaciones.
        """
        # Clonar datos para preservar inmutabilidad
        df_actual = df_cartera_actual_raw.copy()
        df_anterior = df_cartera_anterior_raw.copy()
        df_recaudos = df_recaudos_raw.copy()

        # Resolver nombres de columnas normalizadas
        col_id_actual = find_column_name(df_actual, "Nº ident.fis.1", ["Documento", "Identificacion", "NIT", "Cedula"])
        col_pol_actual = find_column_name(df_actual, "Póliza", ["Poliza", "Numero Poliza", "No Poliza"])

        col_id_ant = find_column_name(df_anterior, "Nº ident.fis.1", ["Documento", "Identificacion", "NIT", "Cedula"])
        col_pol_ant = find_column_name(df_anterior, "Póliza", ["Poliza", "Numero Poliza", "No Poliza"])

        col_id_rec = find_column_name(df_recaudos, "Nº ident.fis.1", ["Documento", "Identificacion", "NIT", "Cedula"])
        col_pol_rec = find_column_name(df_recaudos, "Póliza", ["Poliza", "Numero Poliza", "No Poliza"])

        # Formatear columnas críticas a string limpio
        df_actual["_clean_id"] = df_actual[col_id_actual].apply(clean_id_value) if col_id_actual in df_actual.columns else ""
        df_actual["_clean_poliza"] = df_actual[col_pol_actual].apply(clean_policy_value) if col_pol_actual in df_actual.columns else ""

        df_anterior["_clean_id"] = df_anterior[col_id_ant].apply(clean_id_value) if col_id_ant in df_anterior.columns else ""
        df_anterior["_clean_poliza"] = df_anterior[col_pol_ant].apply(clean_policy_value) if col_pol_ant in df_anterior.columns else ""

        df_recaudos["_clean_id"] = df_recaudos[col_id_rec].apply(clean_id_value) if col_id_rec in df_recaudos.columns else ""
        df_recaudos["_clean_poliza"] = df_recaudos[col_pol_rec].apply(clean_policy_value) if col_pol_rec in df_recaudos.columns else ""

        # Columnas operativas
        col_operacion = find_column_name(df_actual, "Operación", ["Operacion", "Tipo Operacion"])
        col_medio_pago = find_column_name(df_actual, "Medio Pago", ["Medio de Pago", "Medio pago", "Forma Pago"])
        col_dias_cartera = find_column_name(df_actual, "Dias Cartera", ["Días Cartera", "Dias mora", "Dias"])
        col_saldo = find_column_name(df_actual, "Saldo a pagar", ["Saldo", "Valor cartera", "Total a pagar"])
        col_tomador = find_column_name(df_actual, "Nombre Tomador", ["Tomador", "Nombre Cliente", "Cliente"])
        col_solucion = find_column_name(df_actual, "Solución", ["Solucion", "Ramo", "Plan", "Producto"])
        col_resp_cobro = find_column_name(df_actual, "Resp.Cobro", ["Resp Cobro", "Respuesta Cobro", "Motivo Rechazo", "Detalle Rechazo"])

        # Valores auxiliares para lógica
        operacion_series = df_actual[col_operacion].astype(str).str.strip().str.lower() if col_operacion in df_actual.columns else pd.Series([""] * len(df_actual))
        medio_pago_series = df_actual[col_medio_pago].astype(str).str.strip().str.lower() if col_medio_pago in df_actual.columns else pd.Series([""] * len(df_actual))
        dias_cartera_num = parse_numeric(df_actual[col_dias_cartera]) if col_dias_cartera in df_actual.columns else pd.Series([0] * len(df_actual))
        saldo_num = parse_numeric(df_actual[col_saldo]) if col_saldo in df_actual.columns else pd.Series([0.0] * len(df_actual))

        # =========================================================================
        # PASO 1: Separar Devoluciones
        # Filtra filas donde Operación == 'Devolución' -> df_devoluciones
        # df_base: exclusivamente donde Operación == 'Cobro'
        # =========================================================================
        mask_devolucion = operacion_series.str.contains("devoluci", na=False)
        df_devoluciones = df_actual[mask_devolucion].copy()

        mask_cobro = operacion_series.str.contains("cobro", na=False)
        # Si por alguna razón ninguna fila dice 'cobro', tomamos todo lo que no sea devolución para no vaciar por error
        if not mask_cobro.any() and not df_actual.empty:
            df_base = df_actual[~mask_devolucion].copy()
        else:
            df_base = df_actual[mask_cobro].copy()

        total_actual_cargadas = len(df_actual)
        total_devoluciones = len(df_devoluciones)

        # Actualizar series auxiliares alineadas a df_base
        base_medio_pago = df_base[col_medio_pago].astype(str).str.strip().str.lower() if col_medio_pago in df_base.columns else pd.Series([""] * len(df_base), index=df_base.index)
        base_dias_cartera = parse_numeric(df_base[col_dias_cartera]) if col_dias_cartera in df_base.columns else pd.Series([0] * len(df_base), index=df_base.index)
        base_saldo = parse_numeric(df_base[col_saldo]) if col_saldo in df_base.columns else pd.Series([0.0] * len(df_base), index=df_base.index)

        # =========================================================================
        # PASO 2: Excluir Débitos Prematuros
        # Medio Pago == 'Débito' y Dias Cartera <= 16 -> Eliminar del análisis
        # =========================================================================
        mask_debito = base_medio_pago.str.contains("debit|débit", na=False)
        mask_prematuro = mask_debito & (base_dias_cartera <= DIAS_CARTERA_PREMATURO_MAX)

        total_prematuros_excluidos = int(mask_prematuro.sum())
        # Excluir de df_base
        df_base = df_base[~mask_prematuro].copy()

        # Recalcular series auxiliares para df_base tras exclusión
        base_medio_pago = df_base[col_medio_pago].astype(str).str.strip().str.lower() if col_medio_pago in df_base.columns else pd.Series([""] * len(df_base), index=df_base.index)
        base_dias_cartera = parse_numeric(df_base[col_dias_cartera]) if col_dias_cartera in df_base.columns else pd.Series([0] * len(df_base), index=df_base.index)
        base_saldo = parse_numeric(df_base[col_saldo]) if col_saldo in df_base.columns else pd.Series([0.0] * len(df_base), index=df_base.index)

        # =========================================================================
        # PASO 3: Separar Débitos Fallidos y Crear Mensaje
        # Medio Pago == 'Débito' y Dias Cartera > 16 -> Mover a df_debitos_fallidos
        # REGLA: Excluir los que digan 'Débito exitoso' (pues a ellos no les falló)
        # =========================================================================
        base_resp_cobro = df_base[col_resp_cobro].astype(str).str.strip().str.lower() if col_resp_cobro in df_base.columns else pd.Series([""] * len(df_base), index=df_base.index)
        mask_debito_rem = base_medio_pago.str.contains("debit|débit", na=False)

        # Detectar débitos que resultaron exitosos
        mask_debito_exitoso = mask_debito_rem & base_resp_cobro.str.contains(r"exitos|éxitos|debito exitoso|débito exitoso|^ok$", na=False)
        total_debitos_exitosos_excluidos = int(mask_debito_exitoso.sum())

        # Débitos fallidos reales (mora > 16 y sin éxito reportado)
        mask_fallidos = mask_debito_rem & (base_dias_cartera > DIAS_CARTERA_PREMATURO_MAX) & (~mask_debito_exitoso)

        df_debitos_fallidos = df_base[mask_fallidos].copy()
        # Retirar de df_base tanto los fallidos como los exitosos (los exitosos ya se cobraron)
        df_base = df_base[~(mask_fallidos | mask_debito_exitoso)].copy()

        # Generar Mensaje_Contacto en df_debitos_fallidos tomando SOLO el primer nombre
        def _generar_mensaje(row):
            primer_nombre = extract_first_name(row.get(col_tomador, ""))
            solucion = str(row.get(col_solucion, "")).strip() if col_solucion in row and pd.notna(row.get(col_solucion)) else "seguros"
            poliza = str(row.get(col_pol_actual, "")).strip() if col_pol_actual in row and pd.notna(row.get(col_pol_actual)) else ""
            if not poliza and "_clean_poliza" in row:
                poliza = str(row.get("_clean_poliza", ""))

            resp_val = row.get(col_resp_cobro, None) if col_resp_cobro in row else None
            if pd.isna(resp_val) or resp_val is None or str(resp_val).strip() == "" or str(resp_val).lower() in ("nan", "none", "null"):
                resp_str = "motivos no especificados"
            else:
                resp_str = str(resp_val).strip()

            return MENSAJE_DEBITO_FALLIDO_TEMPLATE.format(
                primer_nombre=primer_nombre,
                solucion=solucion,
                poliza=poliza,
                resp_cobro=resp_str,
            )

        if not df_debitos_fallidos.empty:
            df_debitos_fallidos["Mensaje_Contacto"] = df_debitos_fallidos.apply(_generar_mensaje, axis=1)
        else:
            df_debitos_fallidos["Mensaje_Contacto"] = pd.Series(dtype=str)

        # =========================================================================
        # PASO 4: Clientes de Alto Impacto (Urgentes)
        # REGLA ESTRICTA: Solo clientes con Dias Cartera > 1 (no negativos ni 0 ni 1)
        # Extraer hacia df_alto_impacto si cumple morosidad positiva (>1) Y cualquiera:
        # a) Nº ident.fis.1 cruza positivamente con clientes_elite.csv
        # b) Saldo a pagar > 3,000,000
        # c) Dias Cartera >= 45
        # =========================================================================
        elite_set = self.elite_manager.load_elite_set()

        base_dias_cartera = parse_numeric(df_base[col_dias_cartera]) if col_dias_cartera in df_base.columns else pd.Series([0] * len(df_base), index=df_base.index)
        base_saldo = parse_numeric(df_base[col_saldo]) if col_saldo in df_base.columns else pd.Series([0.0] * len(df_base), index=df_base.index)
        base_clean_id = df_base["_clean_id"]

        cond_mora_positiva = base_dias_cartera > 1
        cond_elite = base_clean_id.isin(elite_set)
        cond_saldo = base_saldo > SALDO_ALTO_IMPACTO_MIN
        cond_dias = base_dias_cartera >= DIAS_CARTERA_ALTO_IMPACTO_MIN

        mask_alto_impacto = cond_mora_positiva & (cond_elite | cond_saldo | cond_dias)

        df_alto_impacto = df_base[mask_alto_impacto].copy()
        # Retirar de df_base
        df_base = df_base[~mask_alto_impacto].copy()

        # Añadir columna descriptiva de motivo de impacto (gran valor para el usuario operativo)
        def _motivo_impacto(idx):
            motivos = []
            if cond_elite.loc[idx]:
                motivos.append("Cliente Élite")
            if cond_saldo.loc[idx]:
                motivos.append("Saldo > $3.000.000")
            if cond_dias.loc[idx]:
                motivos.append("Mora >= 45 días")
            return " + ".join(motivos)

        if not df_alto_impacto.empty:
            df_alto_impacto["Motivo_Alto_Impacto"] = [ _motivo_impacto(idx) for idx in df_alto_impacto.index ]
        else:
            df_alto_impacto["Motivo_Alto_Impacto"] = pd.Series(dtype=str)

        # =========================================================================
        # PASO 5: Pólizas a Revisar por Posible Cancelación
        # Pólizas de Cartera Semana Anterior que NO existen en Cartera Actual
        # y TAMPOCO existen en Recaudos.
        # =========================================================================
        polizas_actual_set = set(df_actual["_clean_poliza"].dropna()) - {""}
        polizas_recaudo_set = set(df_recaudos["_clean_poliza"].dropna()) - {""}

        mask_cancelacion = (~df_anterior["_clean_poliza"].isin(polizas_actual_set)) & (~df_anterior["_clean_poliza"].isin(polizas_recaudo_set)) & (df_anterior["_clean_poliza"] != "")

        df_revisar_cancelacion = df_anterior[mask_cancelacion].copy()

        # =========================================================================
        # PASO 6: Validación de No Contactar (Google Sheets)
        # Cruzar lista de excluidos con Nº ident.fis.1 de TODOS los DataFrames
        # Añadir columna 'Estado_Contacto' = 'No contactar' si cruza positivamente
        # =========================================================================
        no_contactar_set, fuente_no_contactar, status_no_contactar = self.no_contactar_service.fetch_no_contactar(
            manual_file=manual_no_contactar_file
        )

        all_dfs = {
            "base": df_base,
            "devoluciones": df_devoluciones,
            "debitos_fallidos": df_debitos_fallidos,
            "alto_impacto": df_alto_impacto,
            "revisar_cancelacion": df_revisar_cancelacion,
        }

        # Asignar Estado_Contacto y limpiar columnas auxiliares internas
        for key, df in all_dfs.items():
            if "_clean_id" in df.columns:
                mask_no_contact = df["_clean_id"].isin(no_contactar_set)
                df["Estado_Contacto"] = ""
                df.loc[mask_no_contact, "Estado_Contacto"] = "No contactar"
            else:
                df["Estado_Contacto"] = ""

            # Eliminar columnas auxiliares que creamos internamente
            df.drop(columns=["_clean_id", "_clean_poliza"], errors="ignore", inplace=True)

        return {
            "dfs": {
                "Base depurada": all_dfs["base"],
                "Devoluciones": all_dfs["devoluciones"],
                "Débitos fallidos": all_dfs["debitos_fallidos"],
                "Alto impacto": all_dfs["alto_impacto"],
                "Revisar cancelación": all_dfs["revisar_cancelacion"],
            },
            "metrics": {
                "total_actual": total_actual_cargadas,
                "devoluciones": len(all_dfs["devoluciones"]),
                "prematuros_excluidos": total_prematuros_excluidos,
                "debitos_exitosos_excluidos": total_debitos_exitosos_excluidos,
                "debitos_fallidos": len(all_dfs["debitos_fallidos"]),
                "alto_impacto": len(all_dfs["alto_impacto"]),
                "base_depurada": len(all_dfs["base"]),
                "revisar_cancelacion": len(all_dfs["revisar_cancelacion"]),
                "total_no_contactar": sum((df["Estado_Contacto"] == "No contactar").sum() for df in all_dfs.values()),
                "total_elite_registrados": len(elite_set),
            },
            "no_contactar_info": {
                "fuente": fuente_no_contactar,
                "estado": status_no_contactar,
                "total_excluidos": len(no_contactar_set),
            },
        }
