import io
import re
import urllib.parse
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd
import requests

from config import (
    ASESORES_CACHE_PATH,
    CLIENTES_ELITE_PATH,
    DIAS_CARTERA_ALTO_IMPACTO_MIN,
    DIAS_CARTERA_PREMATURO_MAX,
    GOOGLE_SHEETS_ASESORES_URL,
    GOOGLE_SHEETS_ASESORES_VIEW_URL,
    GOOGLE_SHEETS_NO_CONTACTAR_URL,
    GOOGLE_SHEETS_NO_CONTACTAR_VIEW_URL,
    MENSAJE_DEBITO_FALLIDO_TEMPLATE,
    NO_CONTACTAR_CACHE_PATH,
    SALDO_ALTO_IMPACTO_MIN,
    SALDO_ALTO_IMPACTO_MENSUAL_MIN,
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
        elif "/1900" in s or "1900-" in s or "/1904" in s:
            # Manejo de celdas formateadas accidentalmente como fecha de 1900 en Excel (ej. 25/02/1900 -> 56 días)
            try:
                dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
                if pd.notna(dt):
                    excel_base = pd.Timestamp("1899-12-30")
                    days_diff = (dt - excel_base).days
                    if 0 < days_diff < 2000:
                        return float(days_diff)
            except Exception:
                pass

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
            df_default = pd.DataFrame(columns=["Nº ident.fis.1", "Nombre Tomador", "Segmento"])
            df_default.to_csv(self.file_path, index=False, encoding="utf-8")

    def get_dataframe(self) -> pd.DataFrame:
        """Carga el DataFrame de clientes élite para edición interactiva."""
        self.ensure_default_file()
        try:
            df = pd.read_csv(self.file_path, dtype=str).fillna("")
            col_id = find_column_name(df, "Nº ident.fis.1", ["Documento", "NIT", "Identificacion", "Cedula", "DNI", "DNI Tomador"])
            if col_id != "Nº ident.fis.1" and col_id in df.columns:
                df.rename(columns={col_id: "Nº ident.fis.1"}, inplace=True)
            if "Nº ident.fis.1" not in df.columns:
                df["Nº ident.fis.1"] = ""
            if "Nombre Tomador" not in df.columns:
                df["Nombre Tomador"] = ""
            if "Segmento" not in df.columns:
                df["Segmento"] = "Élite"
            return df[["Nº ident.fis.1", "Nombre Tomador", "Segmento"]]
        except Exception:
            return pd.DataFrame(columns=["Nº ident.fis.1", "Nombre Tomador", "Segmento"])

    def save_dataframe(self, df: pd.DataFrame) -> Tuple[bool, str, int]:
        """Guarda un DataFrame directamente en clientes_elite.csv, normalizando cédulas."""
        try:
            df_to_save = df.copy()
            col_id = "Nº ident.fis.1" if "Nº ident.fis.1" in df_to_save.columns else df_to_save.columns[0]
            # Descartar filas donde la identificación esté en blanco
            df_to_save = df_to_save[df_to_save[col_id].astype(str).str.strip() != ""]
            # Limpiar documento
            df_to_save[col_id] = df_to_save[col_id].apply(clean_id_value)
            # Eliminar duplicados
            df_to_save.drop_duplicates(subset=[col_id], inplace=True)

            if "Nombre Tomador" not in df_to_save.columns:
                df_to_save["Nombre Tomador"] = ""
            if "Segmento" not in df_to_save.columns:
                df_to_save["Segmento"] = "Élite"
            else:
                df_to_save["Segmento"] = df_to_save["Segmento"].replace("", "Élite").fillna("Élite")

            self.file_path.parent.mkdir(parents=True, exist_ok=True)
            df_to_save.to_csv(self.file_path, index=False, encoding="utf-8")
            return True, f"Listado guardado exitosamente con {len(df_to_save)} clientes élite.", len(df_to_save)
        except Exception as e:
            return False, f"Error al guardar clientes élite: {str(e)}", 0

    def add_client(self, dni: str, nombre: str = "", segmento: str = "Élite") -> Tuple[bool, str]:
        """Añade un cliente individual a la lista permanente."""
        clean_dni = clean_id_value(dni)
        if not clean_dni:
            return False, "La identificación suministrada no es válida."
        df = self.get_dataframe()
        if clean_dni in set(df["Nº ident.fis.1"].apply(clean_id_value)):
            return False, f"El documento {clean_dni} ya se encuentra registrado como cliente élite."
        
        new_row = pd.DataFrame([{
            "Nº ident.fis.1": clean_dni,
            "Nombre Tomador": nombre.strip().upper(),
            "Segmento": segmento.strip() or "Élite"
        }])
        df_updated = pd.concat([df, new_row], ignore_index=True)
        success, msg, _ = self.save_dataframe(df_updated)
        return success, f"Cliente {clean_dni} ({nombre}) agregado exitosamente." if success else msg

    def remove_client(self, dni: str) -> Tuple[bool, str]:
        """Elimina un cliente de la lista permanente por su identificación."""
        clean_dni = clean_id_value(dni)
        df = self.get_dataframe()
        before_count = len(df)
        df_filtered = df[df["Nº ident.fis.1"].apply(clean_id_value) != clean_dni]
        if len(df_filtered) == before_count:
            return False, f"No se encontró ningún cliente con la identificación {clean_dni}."
        success, msg, _ = self.save_dataframe(df_filtered)
        return success, f"Cliente {clean_dni} eliminado de clientes élite." if success else msg

    def load_elite_set(self) -> Set[str]:
        """Carga el conjunto de documentos élite como strings limpios."""
        self.ensure_default_file()
        try:
            df = pd.read_csv(self.file_path, dtype=str)
            col_id = find_column_name(df, "Nº ident.fis.1", ["Documento", "NIT", "Identificacion", "Cedula"])
            if col_id in df.columns:
                return {clean_id_value(v) for v in df[col_id] if clean_id_value(v)}
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
            series_id = df[col_id] if col_id in df.columns else df.iloc[:, 0]

            cleaned_ids = [clean_id_value(x) for x in series_id if clean_id_value(x)]
            unique_ids = sorted(list(set(cleaned_ids)))

            if not unique_ids:
                return False, "No se encontraron identificaciones válidas en el archivo.", 0

            df_to_save = pd.DataFrame({"Nº ident.fis.1": unique_ids, "Nombre Tomador": "", "Segmento": "Élite"})
            return self.save_dataframe(df_to_save)
        except Exception as e:
            return False, f"Error al procesar el archivo de clientes élite: {str(e)}", 0


class NoContactarService:
    """Gestiona la obtención de clientes en lista de No Contactar directamente desde Google Sheets."""

    def __init__(
        self,
        url: str = GOOGLE_SHEETS_NO_CONTACTAR_URL,
        view_url: str = GOOGLE_SHEETS_NO_CONTACTAR_VIEW_URL,
        cache_path: Path = NO_CONTACTAR_CACHE_PATH,
    ):
        self.url = url
        self.view_url = view_url
        self.cache_path = Path(cache_path)
        self.ensure_default_file()

    def ensure_default_file(self):
        """Asegura que el archivo de caché local exista."""
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        if not self.cache_path.exists():
            df_default = pd.DataFrame(columns=["Nº ident.fis.1", "Motivo"])
            df_default.to_csv(self.cache_path, index=False, encoding="utf-8")

    def count(self) -> int:
        """Devuelve el total de pólizas/registros a no contactar registrados en la caché local."""
        self.ensure_default_file()
        try:
            df = pd.read_csv(self.cache_path, dtype=str).fillna("")
            return len(df)
        except Exception:
            return 0

    def fetch_dataframe(self) -> Tuple[pd.DataFrame, str, str]:
        """
        Obtiene la tabla completa de No Contactar directamente desde Google Sheets (pestaña Cartera_Excluidos).
        Si la conexión falla o el acceso es restringido, recurre a la caché local.
        Retorna: (df, fuente, mensaje_estado)
        """
        warning_msg = ""
        # 1. Intentar consultar Google Sheets en vivo
        try:
            import urllib3
            urllib3.disable_warnings()

            resp = requests.get(
                self.url,
                timeout=10,
                headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"},
                verify=False,
            )
            content_type = resp.headers.get("Content-Type", "").lower()
            is_html = "text/html" in content_type or "<!doctype html" in resp.text[:200].lower()

            if resp.status_code == 200 and not is_html:
                df_sheet = pd.read_csv(io.StringIO(resp.text), dtype=str).fillna("")
                # Encontrar la columna de documento
                col_id = find_column_name(
                    df_sheet,
                    "Nº ident",
                    ["Nº ident.fis.1", "Documento", "NIT", "Identificacion", "Cedula", "DNI", "No ident"],
                )
                if col_id in df_sheet.columns:
                    # Crear columna normalizada para cruce seguro sin alterar el texto original
                    df_sheet["_id_clean"] = df_sheet[col_id].apply(clean_id_value)
                    # Descartar filas totalmente vacías
                    df_sheet = df_sheet[df_sheet["_id_clean"] != ""]

                if not df_sheet.empty:
                    # Respaldar en caché local
                    self.cache_path.parent.mkdir(parents=True, exist_ok=True)
                    df_sheet.to_csv(self.cache_path, index=False, encoding="utf-8")
                    return df_sheet, "Google Sheets (Sincronización en vivo)", "OK"
                else:
                    warning_msg = "La hoja de Google Sheets se leyó correctamente pero no contiene registros."

            elif resp.status_code == 401 or is_html:
                warning_msg = (
                    "El archivo de Google Sheets está en modo 'Restringido' en Google Drive (HTTP 401). "
                    "Para sincronización directa sin pedir cuenta de Google, cámbialo a 'Cualquier persona con el enlace' (Lector)."
                )
            else:
                warning_msg = f"Google Sheets respondió con código de estado HTTP {resp.status_code}."
        except Exception as e:
            warning_msg = f"No se pudo conectar a Google Sheets ({type(e).__name__})."

        # 2. Respaldo en caché local
        if self.cache_path.exists():
            try:
                df_cache = pd.read_csv(self.cache_path, dtype=str).fillna("")
                col_id = find_column_name(
                    df_cache,
                    "Nº ident",
                    ["Nº ident.fis.1", "Documento", "NIT", "Identificacion", "Cedula", "DNI", "_id_clean"],
                )
                if col_id in df_cache.columns:
                    df_cache["_id_clean"] = df_cache[col_id].apply(clean_id_value)
                    df_cache = df_cache[df_cache["_id_clean"] != ""]

                if not df_cache.empty:
                    return (
                        df_cache,
                        "Caché local de respaldo",
                        f"{warning_msg} Mostrando los {len(df_cache)} registros de la última copia local.",
                    )
            except Exception:
                pass

        return pd.DataFrame(), "Sin datos", f"{warning_msg} No se pudieron cargar datos."

    def fetch_no_contactar(self) -> Tuple[Set[str], str, str]:
        """
        Obtiene el conjunto de identificaciones únicas a no contactar.
        Retorna: (set_de_ids, fuente_utilizada, estado_o_advertencia)
        """
        df, fuente, estado = self.fetch_dataframe()
        if not df.empty:
            if "_id_clean" in df.columns:
                ids = set(df["_id_clean"].tolist())
                return ids, fuente, estado
            col_id = find_column_name(
                df,
                "Nº ident",
                ["Nº ident.fis.1", "Documento", "NIT", "Identificacion", "Cedula", "DNI"],
            )
            if col_id in df.columns:
                ids = {clean_id_value(x) for x in df[col_id] if clean_id_value(x)}
                return ids, fuente, estado
        return set(), fuente, estado


class AsesoresService:
    """Gestiona el mapeo de Asesor Principal por número de póliza directamente desde Google Sheets."""

    def __init__(
        self,
        url: str = GOOGLE_SHEETS_ASESORES_URL,
        view_url: str = GOOGLE_SHEETS_ASESORES_VIEW_URL,
        cache_path: Path = ASESORES_CACHE_PATH,
    ):
        self.url = url
        self.view_url = view_url
        self.cache_path = Path(cache_path)
        self.ensure_default_file()

    def ensure_default_file(self):
        """Asegura que el archivo de caché local exista."""
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        if not self.cache_path.exists():
            df_default = pd.DataFrame(columns=["Número de póliza", "Asesor principal"])
            df_default.to_csv(self.cache_path, index=False, encoding="utf-8")

    def fetch_mapping(self) -> Tuple[Dict[str, str], str, str]:
        """
        Obtiene el diccionario de mapeo {poliza_limpia: asesor_principal}.
        Retorna: (dict_map, fuente, estado)
        """
        warning_msg = ""
        # 1. Intentar consultar Google Sheets en vivo
        try:
            import urllib3
            urllib3.disable_warnings()

            resp = requests.get(
                self.url,
                timeout=12,
                headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"},
                verify=False,
            )
            content_type = resp.headers.get("Content-Type", "").lower()
            is_html = "text/html" in content_type or "<!doctype html" in resp.text[:200].lower()

            if resp.status_code == 200 and not is_html:
                df_sheet = pd.read_csv(io.StringIO(resp.text), dtype=str).fillna("")
                col_pol = find_column_name(
                    df_sheet, "Número de póliza", ["Numero de poliza", "Póliza", "Poliza", "No Poliza"]
                )
                col_ase = find_column_name(
                    df_sheet, "Asesor principal", ["Asesor", "Comercial", "Asesor Principal"]
                )

                if col_pol in df_sheet.columns and col_ase in df_sheet.columns:
                    # Guardar respaldo
                    self.cache_path.parent.mkdir(parents=True, exist_ok=True)
                    df_to_cache = df_sheet[[col_pol, col_ase]].copy()
                    df_to_cache.columns = ["Número de póliza", "Asesor principal"]
                    df_to_cache.to_csv(self.cache_path, index=False, encoding="utf-8")

                    mapping = self._build_mapping_from_df(df_sheet, col_pol, col_ase)
                    return mapping, "Google Sheets (Sincronización en vivo)", "OK"
                else:
                    warning_msg = f"Columnas no encontradas en Google Sheets ({list(df_sheet.columns[:4])})."
            else:
                warning_msg = f"Google Sheets respondió con código HTTP {resp.status_code}."
        except Exception as e:
            warning_msg = f"No se pudo conectar a Google Sheets ({type(e).__name__})."

        # 2. Respaldo en caché local
        if self.cache_path.exists():
            try:
                df_cache = pd.read_csv(self.cache_path, dtype=str).fillna("")
                col_pol = find_column_name(df_cache, "Número de póliza", ["Póliza", "Poliza"])
                col_ase = find_column_name(df_cache, "Asesor principal", ["Asesor"])
                if col_pol in df_cache.columns and col_ase in df_cache.columns:
                    mapping = self._build_mapping_from_df(df_cache, col_pol, col_ase)
                    return mapping, "Caché local de respaldo", f"{warning_msg} Usando copia local ({len(mapping)} pólizas)."
            except Exception:
                pass

        return {}, "Sin datos", f"{warning_msg} No se pudieron cargar asesores."

    def _build_mapping_from_df(self, df: pd.DataFrame, col_pol: str, col_ase: str) -> Dict[str, str]:
        mapping = {}
        for _, row in df.iterrows():
            pol_raw = str(row.get(col_pol, "")).strip()
            ase_raw = str(row.get(col_ase, "")).strip()
            if not pol_raw or pol_raw.lower() in ("nan", "none", "<na>"):
                continue
            if not ase_raw or ase_raw.lower() in ("nan", "none", "<na>"):
                ase_raw = "Por Asignar"

            # Normalizar nombre del asesor si viene con codificación rota
            if "gu" in ase_raw.lower() and "ab" in ase_raw.lower():
                ase_raw = "GuíaBien"

            clean_p = clean_policy_value(pol_raw)
            clean_no_lead = clean_p.lstrip("0")
            if clean_p:
                mapping[clean_p] = ase_raw
            if clean_no_lead:
                mapping[clean_no_lead] = ase_raw
        return mapping


class CarteraPipeline:
    """Ejecutor del flujo de depuración y lógica de negocio de cartera."""

    def __init__(
        self,
        elite_manager: Optional[ClientesEliteManager] = None,
        no_contactar_service: Optional[NoContactarService] = None,
        asesores_service: Optional[AsesoresService] = None,
    ):
        self.elite_manager = elite_manager or ClientesEliteManager()
        self.no_contactar_service = no_contactar_service or NoContactarService()
        self.asesores_service = asesores_service or AsesoresService()

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

        # Mapeo de Asesor Principal por número de póliza desde Google Sheets
        asesores_map, fuente_asesores, status_asesores = self.asesores_service.fetch_mapping()

        def _get_asesor(pol_val):
            cp = clean_policy_value(pol_val)
            if cp in asesores_map:
                return asesores_map[cp]
            cp_no_zero = cp.lstrip("0")
            if cp_no_zero in asesores_map:
                return asesores_map[cp_no_zero]
            return "Por Asignar"

        if col_pol_actual in df_actual.columns:
            df_actual["Asesor principal"] = df_actual[col_pol_actual].apply(_get_asesor)
        else:
            df_actual["Asesor principal"] = "Por Asignar"

        # Columnas operativas
        col_operacion = find_column_name(df_actual, "Operación", ["Operacion", "Tipo Operacion", "Tipo operacion"])
        col_medio_pago = find_column_name(df_actual, "Medio Pago", ["Medio de Pago", "Medio pago", "Forma Pago", "Forma pago", "Canal de pago"])
        col_dias_cartera = find_column_name(df_actual, "Dias Cartera", ["Días Cartera", "Dias mora", "Días mora", "Dias", "Días", "Mora"])
        col_saldo = find_column_name(df_actual, "Saldo a pagar", ["Saldo", "Saldo actual", "Total a pagar", "Total", "Valor cartera", "Valor a pagar", "Valor prima", "Valor saldo", "Valor"])
        col_tomador = find_column_name(df_actual, "Nombre Tomador", ["Tomador", "Nombre Cliente", "Cliente"])
        col_solucion = find_column_name(df_actual, "Solución", ["Solucion", "Ramo", "Plan", "Producto"])
        col_resp_cobro = find_column_name(df_actual, "Resp.Cobro", ["Resp Cobro", "Respuesta Cobro", "Motivo Rechazo", "Detalle Rechazo"])
        col_periodicidad = find_column_name(
            df_actual,
            "Perioricidad",
            [
                "Periodicidad", "Periodicidad pago", "Periodicidad de pago",
                "Periodo", "Periodo pago", "Periodo de pago",
                "Frecuencia", "Frecuencia Pago", "Frecuencia de pago",
                "Vigencia pago", "Forma de pago", "Forma Pago"
            ]
        )

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
        base_periodicidad = df_base[col_periodicidad].astype(str).str.strip().str.lower() if col_periodicidad in df_base.columns else pd.Series([""] * len(df_base), index=df_base.index)
        base_clean_id = df_base["_clean_id"]

        # Función auxiliar de generación de mensaje para contacto
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

        # =========================================================================
        # PASO 2: Clientes de Alto Impacto (Prioridad Comercial y Financiera)
        # REGLA ESTRICTA: Pólizas con Dias Cartera estrictamente positivos (> 0).
        # Criterios:
        # a) Nº ident.fis.1 cruza positivamente con clientes_elite.csv
        # b) Saldo a pagar > 3,000,000
        # c) Dias Cartera >= 45
        # d) Periodicidad mensual Y Saldo a pagar > 1,000,000
        # NOTA CLAVE: Alto Impacto tiene prioridad ejecutiva para que pólizas críticas
        # no sean descartadas como débitos prematuros ni se confundan con débitos ordinarios.
        # =========================================================================
        elite_set = self.elite_manager.load_elite_set()

        cond_mora_positiva = base_dias_cartera > 0
        cond_elite = base_clean_id.isin(elite_set)
        cond_saldo = base_saldo > SALDO_ALTO_IMPACTO_MIN
        cond_dias = base_dias_cartera >= DIAS_CARTERA_ALTO_IMPACTO_MIN
        cond_mensual = base_periodicidad.str.contains(r"^m$|mensua|mes", case=False, na=False)
        cond_mensual_saldo = cond_mensual & (base_saldo > SALDO_ALTO_IMPACTO_MENSUAL_MIN)

        mask_alto_impacto = cond_mora_positiva & (cond_elite | cond_saldo | cond_dias | cond_mensual_saldo)

        df_alto_impacto = df_base[mask_alto_impacto].copy()
        # Retirar de df_base para que no se dupliquen
        df_base = df_base[~mask_alto_impacto].copy()

        # Añadir columna descriptiva de motivo de impacto
        def _motivo_impacto(idx):
            motivos = []
            if cond_elite.loc[idx]:
                motivos.append("Cliente Élite")
            if cond_saldo.loc[idx]:
                motivos.append("Saldo > $3.000.000")
            if cond_dias.loc[idx]:
                motivos.append("Mora >= 45 días")
            if cond_mensual_saldo.loc[idx]:
                motivos.append("Mensual con saldo > $1.000.000")
            return " + ".join(motivos)

        if not df_alto_impacto.empty:
            df_alto_impacto["Motivo_Alto_Impacto"] = [ _motivo_impacto(idx) for idx in df_alto_impacto.index ]

            def _generar_mensaje_si_debito_fallido(row):
                medio = str(row.get(col_medio_pago, "")).strip().lower()
                # Solo aplica mensaje de débito si el medio de pago es efectivamente Débito
                if not any(k in medio for k in ["debit", "débit"]):
                    return ""

                resp_val = row.get(col_resp_cobro, None) if col_resp_cobro in row else None
                resp_str = "" if (pd.isna(resp_val) or resp_val is None or str(resp_val).lower() in ("nan", "none", "null", "<na>")) else str(resp_val).strip()

                # Si es débito exitoso u ok, no tiene rechazo
                if re.search(r"exitos|éxitos|debito exitoso|débito exitoso|^ok$", resp_str.lower()):
                    return ""

                dias_val = row.get(col_dias_cartera, 0)
                dias_num = parse_numeric(pd.Series([dias_val])).iloc[0]

                # Aplica mensaje si tiene más de 16 días o si ya presenta respuesta de rechazo
                if dias_num > DIAS_CARTERA_PREMATURO_MAX or resp_str != "":
                    return _generar_mensaje(row)

                return ""

            df_alto_impacto["Mensaje_Contacto"] = df_alto_impacto.apply(_generar_mensaje_si_debito_fallido, axis=1)
        else:
            df_alto_impacto["Motivo_Alto_Impacto"] = pd.Series(dtype=str)
            df_alto_impacto["Mensaje_Contacto"] = pd.Series(dtype=str)

        # Recalcular series auxiliares para df_base (pólizas ordinarias restantes)
        base_medio_pago = df_base[col_medio_pago].astype(str).str.strip().str.lower() if col_medio_pago in df_base.columns else pd.Series([""] * len(df_base), index=df_base.index)
        base_dias_cartera = parse_numeric(df_base[col_dias_cartera]) if col_dias_cartera in df_base.columns else pd.Series([0] * len(df_base), index=df_base.index)
        base_saldo = parse_numeric(df_base[col_saldo]) if col_saldo in df_base.columns else pd.Series([0.0] * len(df_base), index=df_base.index)

        # =========================================================================
        # PASO 3: Excluir Débitos Prematuros Ordinarios
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
        base_resp_cobro = df_base[col_resp_cobro].astype(str).str.strip().str.lower() if col_resp_cobro in df_base.columns else pd.Series([""] * len(df_base), index=df_base.index)

        # =========================================================================
        # PASO 4: Separar Débitos Fallidos Ordinarios y Crear Mensaje
        # Medio Pago == 'Débito' y Dias Cartera > 16 -> Mover a df_debitos_fallidos
        # REGLA: Excluir los que digan 'Débito exitoso' (pues ya pagaron)
        # =========================================================================
        mask_debito_rem = base_medio_pago.str.contains("debit|débit", na=False)
        mask_debito_exitoso = mask_debito_rem & base_resp_cobro.str.contains(r"exitos|éxitos|debito exitoso|débito exitoso|^ok$", na=False)
        total_debitos_exitosos_excluidos = int(mask_debito_exitoso.sum())

        mask_fallidos = mask_debito_rem & (base_dias_cartera > DIAS_CARTERA_PREMATURO_MAX) & (~mask_debito_exitoso)

        df_debitos_fallidos = df_base[mask_fallidos].copy()
        # Retirar de df_base tanto los fallidos como los exitosos
        df_base = df_base[~(mask_fallidos | mask_debito_exitoso)].copy()

        if not df_debitos_fallidos.empty:
            df_debitos_fallidos["Mensaje_Contacto"] = df_debitos_fallidos.apply(_generar_mensaje, axis=1)
        else:
            df_debitos_fallidos["Mensaje_Contacto"] = pd.Series(dtype=str)

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
        # Sincronización directa desde Google Sheets
        # Cruzar lista de excluidos con Nº ident.fis.1 de TODOS los DataFrames
        # Añadir columna 'Estado_Contacto' = 'No contactar' si cruza positivamente
        # =========================================================================
        no_contactar_set, fuente_no_contactar, status_no_contactar = self.no_contactar_service.fetch_no_contactar()

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

        # =========================================================================
        # PASO 7: Construcción del Resumen Ejecutivo por Comercial y Hojas por Asesor
        # =========================================================================
        df_cobros_todos = df_actual[~mask_devolucion].copy() if not df_actual.empty else pd.DataFrame()
        resumen_rows = []
        if not df_cobros_todos.empty:
            df_cobros_todos["_saldo_num"] = parse_numeric(df_cobros_todos[col_saldo]) if col_saldo in df_cobros_todos.columns else 0.0
            df_cobros_todos["_dias_num"] = parse_numeric(df_cobros_todos[col_dias_cartera]) if col_dias_cartera in df_cobros_todos.columns else 0

            ai_polizas_set = set(df_alto_impacto[col_pol_actual].apply(clean_policy_value)) if not df_alto_impacto.empty and col_pol_actual in df_alto_impacto.columns else set()
            deb_fallidos_set = set(df_debitos_fallidos[col_pol_actual].apply(clean_policy_value)) if not df_debitos_fallidos.empty and col_pol_actual in df_debitos_fallidos.columns else set()

            total_cartera_global = df_cobros_todos["_saldo_num"].sum()

            for asesor, grp in df_cobros_todos.groupby("Asesor principal"):
                grp_polizas = set(grp[col_pol_actual].apply(clean_policy_value))
                casos_ai = len(grp_polizas.intersection(ai_polizas_set))
                casos_deb = len(grp_polizas.intersection(deb_fallidos_set))

                saldo_total = grp["_saldo_num"].sum()
                mask_grp_ai = grp[col_pol_actual].apply(clean_policy_value).isin(ai_polizas_set)
                saldo_ai = grp.loc[mask_grp_ai, "_saldo_num"].sum()
                mora_prom = grp.loc[grp["_dias_num"] > 0, "_dias_num"].mean() if (grp["_dias_num"] > 0).any() else 0.0
                pct_part = (saldo_total / total_cartera_global * 100) if total_cartera_global > 0 else 0.0

                if casos_ai > 0:
                    accion = f"URGENTE: Cobro inmediato de {casos_ai} casos de Alto Impacto (${saldo_ai:,.0f})"
                elif casos_deb > 0:
                    accion = f"ALERTA: Notificar {casos_deb} débitos rechazados por el banco"
                elif mora_prom >= 30:
                    accion = f"SEGUIMIENTO: Cobro preventivo mora antigua (promedio {mora_prom:.0f} días)"
                else:
                    accion = "REGULAR: Seguimiento preventivo habitual"

                resumen_rows.append({
                    "Comercial": asesor,
                    "Total Cartera ($)": saldo_total,
                    "Pólizas en Mora": len(grp),
                    "Casos Alto Impacto": casos_ai,
                    "Saldo Alto Impacto ($)": saldo_ai,
                    "Débitos Fallidos": casos_deb,
                    "Días Mora Promedio": round(mora_prom, 1),
                    "% Cartera": round(pct_part, 1),
                    "Foco de Acción Sugerido": accion,
                })

        df_resumen_comerciales = pd.DataFrame(resumen_rows)
        if not df_resumen_comerciales.empty:
            df_resumen_comerciales.sort_values(by="Total Cartera ($)", ascending=False, inplace=True)

        # Generar hojas individuales por comercial para Alto Impacto
        dfs_por_comercial = {}
        if not df_alto_impacto.empty and "Asesor principal" in df_alto_impacto.columns:
            for asesor in sorted(df_alto_impacto["Asesor principal"].unique()):
                sub_df = df_alto_impacto[df_alto_impacto["Asesor principal"] == asesor].copy()
                clean_name = re.sub(r'[\\/*?:\[\]]', '_', str(asesor))
                sheet_title = f"AI - {clean_name}"[:31]
                dfs_por_comercial[sheet_title] = sub_df

        return {
            "dfs": {
                "Base depurada": all_dfs["base"],
                "Devoluciones": all_dfs["devoluciones"],
                "Débitos fallidos": all_dfs["debitos_fallidos"],
                "Alto impacto": all_dfs["alto_impacto"],
                "Revisar cancelación": all_dfs["revisar_cancelacion"],
            },
            "resumen_comerciales": df_resumen_comerciales,
            "dfs_comerciales": dfs_por_comercial,
            "asesores_info": {
                "fuente": fuente_asesores,
                "estado": status_asesores,
                "total_mapeadas": len(asesores_map),
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
            "diagnostico_alto_impacto": {
                "col_periodicidad": col_periodicidad,
                "col_saldo": col_saldo,
                "col_dias_cartera": col_dias_cartera,
                "total_dias_gt_0": int((parse_numeric(df_actual[col_dias_cartera]) > 0).sum()) if col_dias_cartera in df_actual.columns else 0,
                "total_saldo_gt_1m": int((parse_numeric(df_actual[col_saldo]) > SALDO_ALTO_IMPACTO_MENSUAL_MIN).sum()) if col_saldo in df_actual.columns else 0,
                "total_mensuales": int(df_actual[col_periodicidad].astype(str).str.lower().str.contains(r"^m$|mensua|mes", na=False).sum()) if col_periodicidad in df_actual.columns else 0,
            },
        }
