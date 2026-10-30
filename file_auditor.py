import hashlib
import io
import re
from typing import Dict, List, Optional, Tuple
import pandas as pd


def compute_file_hash(content_bytes: bytes) -> str:
    """Calcula el hash SHA-256 de los bytes de un archivo para detectar duplicados exactos."""
    return hashlib.sha256(content_bytes).hexdigest()


def detect_file_schema(df: pd.DataFrame, filename: str = "") -> Tuple[str, float, List[str]]:
    """
    Determina si un DataFrame corresponde a 'cartera', 'recaudos' o 'desconocido'.
    Basado en la estructura real de informes emitida por Seguros SURA:
    - Cartera contiene exclusivamente columnas de mora/gestión: 'Operación', 'Dias Cartera', 'Saldo a pagar'.
    - Recaudos contiene pagos efectuados: 'Fecha de pago', 'Total', 'Ref/Financiación', o prefijo 'Recaudo'.
    """
    cols_norm = [re.sub(r"[\s\._º°/]", "", str(c)).lower() for c in df.columns]
    fname = filename.lower()

    # Columnas exclusivas de Cartera (nunca aparecen en recaudos de SURA)
    cartera_exclusive = ["operacion", "diascartera", "saldoapagar"]
    cartera_matches = [ind for ind in cartera_exclusive if any(ind in c for c in cols_norm)]

    # Columnas exclusivas o muy típicas de Recaudos
    recaudo_exclusive = ["fechadepago", "fechapago", "fecharecaudo", "reffinanciacion", "valorpagado", "valorrecaudo"]
    recaudo_matches = [ind for ind in recaudo_exclusive if any(ind in c for c in cols_norm)]

    if "recaudo" in fname:
        recaudo_matches.append("nombre_archivo_recaudo")

    # Regla 1: Si tiene columnas o nombre de recaudo y NO tiene 'diascartera' ni 'operacion'
    if recaudo_matches and not cartera_matches:
        return "recaudos", 1.0, recaudo_matches

    # Regla 2: Si tiene columnas exclusivas de cartera
    if cartera_matches:
        return "cartera", 1.0, cartera_matches

    # Regla 3: Si tiene 'Total' y no tiene diascartera ni saldoapagar
    if "total" in cols_norm and not cartera_matches:
        return "recaudos", 0.8, ["total_pago"]

    return "desconocido", 0.0, []


def extract_reference_date(df: pd.DataFrame, schema_type: str = "") -> Tuple[Optional[pd.Timestamp], Optional[str]]:
    """
    Busca columnas de fecha o metadatos para encontrar la fecha de corte o fecha representativa.
    Ignora deliberadamente fechas de vigencia de póliza (ej. Fec.Ini.Vi, Fe.Fin.Vig.) que van
    a años futuros y distorsionan la auditoría temporal.
    """
    ignore_keywords = [
        "fecinivi", "fefinvig", "vigencia", "fechainicio", "fechafin",
        "iniciovig", "finvig", "fechaultimocobro"
    ]

    # Prioridad según tipo de informe
    if schema_type == "recaudos":
        priority_keywords = ["fechadepago", "fechapago", "fecharecaudo", "fecha"]
    else:
        priority_keywords = ["fechacorte", "corte", "fechageneracion", "fechaproceso", "fechavencimiento", "fecha"]

    cols_map = {re.sub(r"[\s\._º°/]", "", str(c)).lower(): c for c in df.columns}

    # 1. Buscar en columnas prioritarias no ignoradas
    for kw in priority_keywords:
        matching_cols = [orig for norm, orig in cols_map.items() if kw in norm and not any(ign in norm for ign in ignore_keywords)]
        for col in matching_cols:
            series = df[col].dropna()
            if series.empty:
                continue
            try:
                dt_series = pd.to_datetime(series, format="mixed", errors="coerce").dropna()
                if dt_series.empty:
                    dt_series = pd.to_datetime(series, dayfirst=True, errors="coerce").dropna()
                if not dt_series.empty:
                    dt_series = dt_series[(dt_series.dt.year >= 2020) & (dt_series.dt.year <= 2040)]
                    if not dt_series.empty:
                        return dt_series.max(), col
            except Exception:
                continue

    # 2. Buscar en cualquier otra fecha no ignorada
    for norm, orig in cols_map.items():
        if any(ign in norm for ign in ignore_keywords):
            continue
        if "fecha" in norm or "fec" in norm:
            series = df[orig].dropna()
            try:
                dt_series = pd.to_datetime(series, format="mixed", errors="coerce").dropna()
                if not dt_series.empty:
                    dt_series = dt_series[(dt_series.dt.year >= 2020) & (dt_series.dt.year <= 2040)]
                    if not dt_series.empty:
                        return dt_series.max(), orig
            except Exception:
                continue

    return None, None


class FileAuditor:
    """Audita y valida la coherencia de los 3 archivos operativos antes del procesamiento."""

    @staticmethod
    def audit_upload(
        file_actual_bytes: bytes,
        file_actual_name: str,
        df_actual: pd.DataFrame,
        file_anterior_bytes: Optional[bytes] = None,
        file_anterior_name: Optional[str] = None,
        df_anterior: Optional[pd.DataFrame] = None,
        file_recaudos_bytes: Optional[bytes] = None,
        file_recaudos_name: Optional[str] = None,
        df_recaudos: Optional[pd.DataFrame] = None,
    ) -> Dict[str, any]:
        """
        Ejecuta todas las reglas de control de calidad y pre-vuelo:
        1. Detección de archivos idénticos (hash).
        2. Detección de esquemas invertidos (recaudo en cartera, etc.).
        3. Detección cronológica (cartera anterior más nueva que la actual).
        4. Frescura de recaudos.
        """
        errors = []
        warnings = []
        info = []
        can_process = True
        auto_swap_recommended = False

        # 1. Hashes y duplicados
        hash_actual = compute_file_hash(file_actual_bytes)
        hash_anterior = compute_file_hash(file_anterior_bytes) if file_anterior_bytes else None
        hash_recaudos = compute_file_hash(file_recaudos_bytes) if file_recaudos_bytes else None

        if hash_anterior and hash_actual == hash_anterior:
            errors.append(
                f"🚨 **Archivos idénticos detectados:** Has cargado exactamente el mismo archivo en 'Cartera Actual' y 'Cartera Anterior' (`{file_actual_name}`). Esto provocaría resultados falsos."
            )
            can_process = False

        if hash_recaudos and hash_actual == hash_recaudos:
            errors.append(
                f"🚨 **Archivos idénticos detectados:** Has cargado el mismo archivo en 'Cartera Actual' y 'Recaudos' (`{file_actual_name}`)."
            )
            can_process = False

        if hash_anterior and hash_recaudos and hash_anterior == hash_recaudos:
            errors.append(
                f"🚨 **Archivos idénticos detectados:** Has cargado el mismo archivo en 'Cartera Anterior' y 'Recaudos' (`{file_anterior_name}`)."
            )
            can_process = False

        # 2. Validación de esquemas
        type_actual, conf_act, _ = detect_file_schema(df_actual, file_actual_name)
        if type_actual == "recaudos":
            errors.append(
                f"⛔ **Esquema incorrecto:** El archivo cargado en 'Cartera Actual' (`{file_actual_name}`) parece ser un informe de **Recaudos**. Por favor revisa la casilla."
            )
            can_process = False

        type_anterior, conf_ant, _ = (
            detect_file_schema(df_anterior, file_anterior_name) if df_anterior is not None else (None, 0, [])
        )
        if type_anterior == "recaudos":
            errors.append(
                f"⛔ **Esquema incorrecto:** El archivo cargado en 'Cartera Anterior' (`{file_anterior_name}`) parece ser un informe de **Recaudos**."
            )
            can_process = False

        type_recaudos, conf_rec, _ = (
            detect_file_schema(df_recaudos, file_recaudos_name) if df_recaudos is not None else (None, 0, [])
        )
        if type_recaudos == "cartera":
            # Verificar si realmente tiene columnas de cobro/mora para evitar falsos positivos
            cols_rec_norm = [re.sub(r"[\s\._º°/]", "", str(c)).lower() for c in df_recaudos.columns]
            if any(c in cols_rec_norm for c in ["diascartera", "operacion"]):
                warnings.append(
                    f"⚠️ **Atención con Recaudos:** El archivo cargado en 'Recaudos' (`{file_recaudos_name}`) tiene estructura de **Cartera** (incluye campos de cobro/mora)."
                )

        # 3. Auditoría Cronológica de Cartera (Actual vs Anterior)
        date_actual, col_date_act = extract_reference_date(df_actual, schema_type="cartera")
        date_anterior, col_date_ant = (
            extract_reference_date(df_anterior, schema_type="cartera") if df_anterior is not None else (None, None)
        )
        date_recaudos, col_date_rec = (
            extract_reference_date(df_recaudos, schema_type="recaudos") if df_recaudos is not None else (None, None)
        )

        cronologia_status = "ok"

        if date_actual and date_anterior:
            if date_anterior > date_actual:
                auto_swap_recommended = True
                cronologia_status = "invertido"
                errors.append(
                    f"🔄 **¡Alerta de Inversión Cronológica!** La Cartera Anterior tiene fecha más reciente (**{date_anterior.strftime('%d/%m/%Y')}**) que la Cartera Actual (**{date_actual.strftime('%d/%m/%Y')}**). Es muy probable que los archivos estén intercambiados."
                )
                can_process = False
            elif date_actual == date_anterior:
                warnings.append(
                    f"⚠️ **Fechas idénticas:** Ambos informes de cartera registran fecha de corte **{date_actual.strftime('%d/%m/%Y')}**. Verifica si corresponden a semanas diferentes."
                )
            else:
                dias_diferencia = (date_actual - date_anterior).days
                info.append(
                    f"📅 **Cronología coherente:** Cartera Anterior ({date_anterior.strftime('%d/%m/%Y')}) ➔ Cartera Actual ({date_actual.strftime('%d/%m/%Y')}) — Diferencia: {dias_diferencia} días."
                )

        # 4. Frescura de Recaudos
        if date_recaudos:
            ref_date = date_actual or pd.Timestamp.now()
            dias_desfase = (ref_date - date_recaudos).days

            if dias_desfase > 35:
                warnings.append(
                    f"⏰ **Recaudo potencialmente desactualizado:** La fecha más reciente en Recaudos es **{date_recaudos.strftime('%d/%m/%Y')}** (hace {dias_desfase} días respecto a la Cartera Actual). Confirma si es el informe de la semana en curso."
                )
            elif dias_desfase < -7:
                warnings.append(
                    f"⚠️ **Recaudo futuro:** Las fechas de recaudo (**{date_recaudos.strftime('%d/%m/%Y')}**) superan la fecha de la cartera actual por más de 7 días."
                )
            else:
                info.append(
                    f"💳 **Recaudos al día:** Fechas registradas hasta **{date_recaudos.strftime('%d/%m/%Y')}** (ventana de {abs(dias_desfase)} días)."
                )

        return {
            "can_process": can_process,
            "errors": errors,
            "warnings": warnings,
            "info": info,
            "auto_swap_recommended": auto_swap_recommended,
            "dates": {
                "actual": date_actual.strftime("%d/%m/%Y") if date_actual else "No detectada",
                "anterior": date_anterior.strftime("%d/%m/%Y") if date_anterior else "No detectada",
                "recaudos": date_recaudos.strftime("%d/%m/%Y") if date_recaudos else "No detectada",
            },
            "types": {
                "actual": type_actual,
                "anterior": type_anterior,
                "recaudos": type_recaudos,
            },
        }

    @staticmethod
    def auto_classify_batch(files_dict: Dict[str, Tuple[bytes, pd.DataFrame]]) -> Dict[str, any]:
        """
        Clasifica automáticamente una colección de archivos subidos en lote:
        Identifica cuál es Recaudos y, entre las carteras, cuál es la Actual (más reciente)
        y cuál es la Anterior (más antigua).
        """
        carteras = []
        recaudos = []
        desconocidos = []

        for name, (b, df) in files_dict.items():
            schema_type, conf, matches = detect_file_schema(df, name)
            date_val, col = extract_reference_date(df, schema_type=schema_type)
            item = {"name": name, "bytes": b, "df": df, "date": date_val, "schema": schema_type, "conf": conf}

            if schema_type == "recaudos":
                recaudos.append(item)
            elif schema_type == "cartera":
                carteras.append(item)
            else:
                desconocidos.append(item)

        # Si tenemos carteras, ordenarlas por fecha (la más reciente es Actual, la anterior es Anterior)
        cartera_actual = None
        cartera_anterior = None

        if len(carteras) >= 2:
            # Ordenar por fecha descendente si hay fechas
            carteras_con_fecha = [c for c in carteras if c["date"] is not None]
            if len(carteras_con_fecha) >= 2:
                carteras.sort(key=lambda x: x["date"] or pd.Timestamp.min, reverse=True)
            cartera_actual = carteras[0]
            cartera_anterior = carteras[1]
        elif len(carteras) == 1:
            cartera_actual = carteras[0]

        recaudo_final = recaudos[0] if recaudos else None

        return {
            "cartera_actual": cartera_actual,
            "cartera_anterior": cartera_anterior,
            "recaudos": recaudo_final,
            "desconocidos": desconocidos,
            "success": cartera_actual is not None and cartera_anterior is not None and recaudo_final is not None,
        }
