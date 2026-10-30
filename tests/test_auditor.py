import sys
from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from file_auditor import FileAuditor, compute_file_hash, detect_file_schema, extract_reference_date


def test_schema_detection():
    df_cartera = pd.DataFrame({
        "Nº ident.fis.1": ["101"],
        "Póliza": ["POL1"],
        "Operación": ["Cobro"],
        "Dias Cartera": [20],
        "Saldo a pagar": [500000],
    })
    tipo, conf, matches = detect_file_schema(df_cartera)
    assert tipo == "cartera"
    assert conf >= 0.5

    df_recaudo = pd.DataFrame({
        "Nº ident.fis.1": ["101"],
        "Póliza": ["POL1"],
        "Valor Recaudo": [200000],
        "Fecha Recaudo": ["2026-10-25"],
    })
    tipo_r, conf_r, matches_r = detect_file_schema(df_recaudo)
    assert tipo_r == "recaudos"


def test_inverted_dates_detection():
    # Cartera Actual con fecha 2026-10-15
    df_act = pd.DataFrame({
        "Nº ident.fis.1": ["101"],
        "Póliza": ["POL1"],
        "Operación": ["Cobro"],
        "Fecha Corte": ["15/10/2026"],
    })
    # Cartera Anterior con fecha 2026-10-25 (MÁS NUEVA! Error de usuario que las invirtió)
    df_ant = pd.DataFrame({
        "Nº ident.fis.1": ["101"],
        "Póliza": ["POL1"],
        "Operación": ["Cobro"],
        "Fecha Corte": ["25/10/2026"],
    })

    audit = FileAuditor.audit_upload(
        file_actual_bytes=b"actual_data",
        file_actual_name="cartera_1.xlsx",
        df_actual=df_act,
        file_anterior_bytes=b"anterior_data",
        file_anterior_name="cartera_2.xlsx",
        df_anterior=df_ant,
    )

    assert audit["can_process"] is False
    assert audit["auto_swap_recommended"] is True
    assert any("Inversión Cronológica" in e for e in audit["errors"])
    print("[PASS] test_inverted_dates_detection passed")


def test_duplicate_files_detection():
    df = pd.DataFrame({"Nº ident.fis.1": ["101"], "Operación": ["Cobro"]})
    audit = FileAuditor.audit_upload(
        file_actual_bytes=b"same_bytes",
        file_actual_name="mismo_archivo.xlsx",
        df_actual=df,
        file_anterior_bytes=b"same_bytes",
        file_anterior_name="mismo_archivo.xlsx",
        df_anterior=df,
    )

    assert audit["can_process"] is False
    assert any("Archivos idénticos" in e for e in audit["errors"])
    print("[PASS] test_duplicate_files_detection passed")


def test_recaudos_in_cartera_detection():
    df_rec = pd.DataFrame({"Valor Recaudo": [100000], "Fecha Recaudo": ["2026-10-20"]})
    audit = FileAuditor.audit_upload(
        file_actual_bytes=b"rec_bytes",
        file_actual_name="recaudos_equivocado.xlsx",
        df_actual=df_rec,
    )

    assert audit["can_process"] is False
    assert any("Esquema incorrecto" in e for e in audit["errors"])
    print("[PASS] test_recaudos_in_cartera_detection passed")


if __name__ == "__main__":
    test_schema_detection()
    test_inverted_dates_detection()
    test_duplicate_files_detection()
    test_recaudos_in_cartera_detection()
    print("\nALL AUDITOR TESTS PASSED SUCCESSFULLY!")
