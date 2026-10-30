import io
import sys
import tempfile
from pathlib import Path

# Agregar raíz al sys.path
BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

import pandas as pd

from cartera_processor import (
    CarteraPipeline,
    ClientesEliteManager,
    NoContactarService,
    clean_id_value,
    extract_first_name,
    parse_numeric,
)
from excel_exporter import export_cartera_to_excel


def test_clean_id_and_numeric():
    assert clean_id_value(" 123456.0 ") == "123456"
    assert clean_id_value("42.680.537") == "42680537"
    assert clean_id_value("NIT 900.244.957-5") == "900244957"
    assert clean_id_value("CC 71234567") == "71234567"
    assert clean_id_value(1017000.0) == "1017000"
    assert clean_id_value(None) == ""
    assert clean_id_value("NaN") == ""

    assert extract_first_name("MARGARITA MARIA TOBON RESTREPO") == "Margarita"
    assert extract_first_name("Dr. Carlos Ruiz") == "Carlos"
    assert extract_first_name("ANA GOMEZ") == "Ana"
    assert extract_first_name(None) == "estimado cliente"

    s = pd.Series(["$ 3.500.000,00", "4,200,000.50", "3000000", "COP 15000", None])
    res = parse_numeric(s)
    assert res.iloc[0] == 3500000.0
    assert res.iloc[1] == 4200000.5
    assert res.iloc[2] == 3000000.0
    assert res.iloc[3] == 15000.0
    assert res.iloc[4] == 0.0
    print("[PASS] test_clean_id_and_numeric passed")


def test_pipeline_complete_flow():
    with tempfile.TemporaryDirectory() as tmp_dir:
        elite_path = Path(tmp_dir) / "clientes_elite.csv"
        elite_df = pd.DataFrame({"Nº ident.fis.1": ["111111", "222222"]})
        elite_df.to_csv(elite_path, index=False)

        elite_mgr = ClientesEliteManager(file_path=elite_path)
        assert elite_mgr.count() == 2

        # Mock service de no contactar con set controlado
        class MockNoContactarService:
            def fetch_no_contactar(self, manual_file=None):
                return {"999999", "111111"}, "Mock Service", "OK"

        class MockAsesoresService:
            def fetch_mapping(self):
                return {
                    "POL-05": "Asesor_NG",
                    "POL-06": "Asesor_NB",
                    "POL-07": "Asesor_NG",
                    "POL-MENSUAL-ALTO": "Asesor_ES",
                }, "Mock Asesores", "OK"

        pipeline = CarteraPipeline(
            elite_manager=elite_mgr,
            no_contactar_service=MockNoContactarService(),
            asesores_service=MockAsesoresService(),
        )

        # 1. Cartera Actual
        data_actual = [
            # 0: Devolución -> Debe ir a df_devoluciones
            {
                "Nº ident.fis.1": "100001",
                "Póliza": "POL-01",
                "Operación": "Devolución",
                "Medio Pago": "Débito",
                "Dias Cartera": "10",
                "Saldo a pagar": "100000",
                "Nombre Tomador": "Carlos Ruiz",
                "Solución": "Autos",
                "Resp.Cobro": "",
            },
            # 1: Débito Prematuro (días <= 16) -> Debe ser eliminado del análisis
            {
                "Nº ident.fis.1": "100002",
                "Póliza": "POL-02",
                "Operación": "Cobro",
                "Medio Pago": "Débito",
                "Dias Cartera": "12",
                "Saldo a pagar": "200000",
                "Nombre Tomador": "Ana Gómez",
                "Solución": "Vida",
                "Resp.Cobro": "",
            },
            # 2: Débito Fallido (días > 16) con Resp.Cobro especificada -> df_debitos_fallidos
            {
                "Nº ident.fis.1": "100003",
                "Póliza": "POL-03",
                "Operación": "Cobro",
                "Medio Pago": "Débito",
                "Dias Cartera": "20",
                "Saldo a pagar": "300000",
                "Nombre Tomador": "Beatriz Mesa Soto",
                "Solución": "Salud",
                "Resp.Cobro": "Fondos insuficientes",
            },
            # 3: Débito Fallido sin Resp.Cobro (debe usar 'motivos no especificados') -> df_debitos_fallidos
            {
                "Nº ident.fis.1": "100004",
                "Póliza": "POL-04",
                "Operación": "Cobro",
                "Medio Pago": "Débito",
                "Dias Cartera": "25",
                "Saldo a pagar": "400000",
                "Nombre Tomador": "David Fernando Soto",
                "Solución": "Hogar",
                "Resp.Cobro": None,
            },
            # 4: Débito Exitoso (días > 16 pero dice 'Débito exitoso') -> NO debe entrar a fallidos!
            {
                "Nº ident.fis.1": "100005",
                "Póliza": "POL-05-EXITO",
                "Operación": "Cobro",
                "Medio Pago": "Débito",
                "Dias Cartera": "22",
                "Saldo a pagar": "350000",
                "Nombre Tomador": "Marcela Restrepo",
                "Solución": "Autos",
                "Resp.Cobro": "Débito exitoso",
            },
            # 5: Alto Impacto por Élite (111111 está en elite Y días > 1) -> df_alto_impacto
            # Además, 111111 está en no contactar!
            {
                "Nº ident.fis.1": "111111",
                "Póliza": "POL-05",
                "Operación": "Cobro",
                "Medio Pago": "Efectivo",
                "Dias Cartera": "10",
                "Saldo a pagar": "500000",
                "Nombre Tomador": "Empresa Elite SAS",
                "Solución": "Pyme",
                "Resp.Cobro": "",
            },
            # 6: Alto Impacto por Saldo (> 3,000,000) Y días > 1 -> df_alto_impacto
            {
                "Nº ident.fis.1": "100006",
                "Póliza": "POL-06",
                "Operación": "Cobro",
                "Medio Pago": "Transferencia",
                "Dias Cartera": "15",
                "Saldo a pagar": "3500000",
                "Nombre Tomador": "Jorge Restrepo",
                "Solución": "Salud Global",
                "Resp.Cobro": "",
            },
            # 7: Alto Impacto por Días (>= 45) -> df_alto_impacto
            {
                "Nº ident.fis.1": "100007",
                "Póliza": "POL-07",
                "Operación": "Cobro",
                "Medio Pago": "Tarjeta Crédito",
                "Dias Cartera": "50",
                "Saldo a pagar": "800000",
                "Nombre Tomador": "Marta Linares",
                "Solución": "Vida",
                "Resp.Cobro": "",
            },
            # 8: Saldo > 3,000,000 PERO Días Cartera = -5 (Negativo!) -> NO DEBE entrar a Alto Impacto! Debe quedarse en Base Depurada
            {
                "Nº ident.fis.1": "100009",
                "Póliza": "POL-NEGATIVO",
                "Operación": "Cobro",
                "Medio Pago": "Transferencia",
                "Dias Cartera": "-5",
                "Saldo a pagar": "4500000",
                "Nombre Tomador": "Cliente Anticipado",
                "Solución": "Salud",
                "Resp.Cobro": "",
            },
            # 8b: Saldo > 3,000,000 PERO Días Cartera = 0 -> NO DEBE entrar a Alto Impacto (solo > 0)! Debe quedarse en Base Depurada
            {
                "Nº ident.fis.1": "100010",
                "Póliza": "POL-MORA-CERO",
                "Operación": "Cobro",
                "Medio Pago": "Transferencia",
                "Dias Cartera": "0",
                "Saldo a pagar": "4000000",
                "Nombre Tomador": "Cliente Al Día",
                "Solución": "Salud",
                "Resp.Cobro": "",
            },
            # 8c: Periodicidad Mensual y Saldo > 1,000,000 con mora > 0 -> DEBE entrar a Alto Impacto! (Transferencia -> sin mensaje)
            {
                "Nº ident.fis.1": "100011",
                "Póliza": "POL-MENSUAL-ALTO",
                "Operación": "Cobro",
                "Medio Pago": "Transferencia",
                "Dias Cartera": "8",
                "Saldo a pagar": "1500000",
                "Perioricidad": "Mensual",
                "Nombre Tomador": "Cliente Mensual Alto",
                "Solución": "Autos",
                "Resp.Cobro": "",
            },
            # 8c2: Alto Impacto por Débito Fallido (> 16 días, saldo > 3M) -> DEBE tener Mensaje_Contacto
            {
                "Nº ident.fis.1": "100013",
                "Póliza": "POL-AI-DEBITO",
                "Operación": "Cobro",
                "Medio Pago": "Débito",
                "Dias Cartera": "22",
                "Saldo a pagar": "3500000",
                "Nombre Tomador": "Carlos Alberto Pérez",
                "Solución": "Vida",
                "Resp.Cobro": "Fondos insuficientes",
            },
            # 8d: Periodicidad Anual y Saldo 1,500,000 (no llega a 3M) -> NO es mensual, se queda en Base Depurada
            {
                "Nº ident.fis.1": "100012",
                "Póliza": "POL-ANUAL-MEDIO",
                "Operación": "Cobro",
                "Medio Pago": "Transferencia",
                "Dias Cartera": "8",
                "Saldo a pagar": "1500000",
                "Perioricidad": "Anual",
                "Nombre Tomador": "Cliente Anual",
                "Solución": "Autos",
                "Resp.Cobro": "",
            },
            # 9: Cobro estándar regular que queda en Base Depurada
            {
                "Nº ident.fis.1": "100008",
                "Póliza": "POL-08",
                "Operación": "Cobro",
                "Medio Pago": "Transferencia",
                "Dias Cartera": "22",
                "Saldo a pagar": "950000",
                "Nombre Tomador": "Pedro Uribe",
                "Solución": "Autos",
                "Resp.Cobro": "",
            },
        ]
        df_actual = pd.DataFrame(data_actual)

        # 2. Cartera Semana Anterior
        data_anterior = [
            {"Nº ident.fis.1": "100008", "Póliza": "POL-08", "Nombre Tomador": "Pedro Uribe"},
            {"Nº ident.fis.1": "100090", "Póliza": "POL-90", "Nombre Tomador": "Laura Rios"},
            {"Nº ident.fis.1": "999999", "Póliza": "POL-99", "Nombre Tomador": "Cliente Cancelado"},
        ]
        df_anterior = pd.DataFrame(data_anterior)

        # 3. Recaudos
        data_recaudos = [
            {"Nº ident.fis.1": "100090", "Póliza": "POL-90", "Total": "600000", "Fecha de pago": "2026-10-25"},
        ]
        df_recaudos = pd.DataFrame(data_recaudos)

        # Ejecutar pipeline
        result = pipeline.process(df_actual, df_anterior, df_recaudos)
        dfs = result["dfs"]

        # Verificaciones
        assert len(dfs["Devoluciones"]) == 1
        assert dfs["Devoluciones"].iloc[0]["Póliza"] == "POL-01"

        # Débitos fallidos (POL-03 y POL-04, POL-05-EXITO EXCLUIDO!)
        assert len(dfs["Débitos fallidos"]) == 2
        polizas_fallidas = set(dfs["Débitos fallidos"]["Póliza"])
        assert "POL-05-EXITO" not in polizas_fallidas
        assert polizas_fallidas == {"POL-03", "POL-04"}

        # Mensaje debe usar solo el PRIMER nombre
        msg_03 = dfs["Débitos fallidos"].loc[dfs["Débitos fallidos"]["Póliza"] == "POL-03", "Mensaje_Contacto"].iloc[0]
        assert "Hola Beatriz," in msg_03
        assert "Beatriz Mesa Soto" not in msg_03

        msg_04 = dfs["Débitos fallidos"].loc[dfs["Débitos fallidos"]["Póliza"] == "POL-04", "Mensaje_Contacto"].iloc[0]
        assert "Hola David," in msg_04
        assert "motivos no especificados" in msg_04

        # Alto Impacto (POL-05: Élite, POL-06: Saldo > 3M, POL-07: Días >= 45, POL-MENSUAL-ALTO: Mensual > 1M, POL-AI-DEBITO: Débito Fallido > 3M)
        # POL-NEGATIVO (-5) y POL-MORA-CERO (0) NO DEBEN estar en Alto Impacto por mora no positiva!
        # POL-ANUAL-MEDIO (1.5M anual) NO debe estar en Alto Impacto por no ser mensual ni > 3M.
        assert len(dfs["Alto impacto"]) == 5
        polizas_impacto = set(dfs["Alto impacto"]["Póliza"])
        assert polizas_impacto == {"POL-05", "POL-06", "POL-07", "POL-MENSUAL-ALTO", "POL-AI-DEBITO"}
        assert "POL-NEGATIVO" not in polizas_impacto
        assert "POL-MORA-CERO" not in polizas_impacto
        assert "POL-ANUAL-MEDIO" not in polizas_impacto

        # Motivo de Alto Impacto descriptivo
        motivo_mensual = dfs["Alto impacto"].loc[dfs["Alto impacto"]["Póliza"] == "POL-MENSUAL-ALTO", "Motivo_Alto_Impacto"].iloc[0]
        assert "Mensual con saldo > $1.000.000" in motivo_mensual

        # REGLA CRÍTICA: Mensaje_Contacto SOLO para los que tienen débito fallido
        msg_mensual = dfs["Alto impacto"].loc[dfs["Alto impacto"]["Póliza"] == "POL-MENSUAL-ALTO", "Mensaje_Contacto"].iloc[0]
        assert msg_mensual == ""  # Es transferencia -> celda en blanco!

        msg_elite = dfs["Alto impacto"].loc[dfs["Alto impacto"]["Póliza"] == "POL-05", "Mensaje_Contacto"].iloc[0]
        assert msg_elite == ""  # Es efectivo -> celda en blanco!

        msg_ai_deb = dfs["Alto impacto"].loc[dfs["Alto impacto"]["Póliza"] == "POL-AI-DEBITO", "Mensaje_Contacto"].iloc[0]
        assert "Hola Carlos," in msg_ai_deb  # Es débito fallido -> SÍ tiene mensaje!
        assert "Fondos insuficientes" in msg_ai_deb

        # Base Depurada: Contiene POL-08, POL-NEGATIVO, POL-MORA-CERO y POL-ANUAL-MEDIO
        polizas_base = set(dfs["Base depurada"]["Póliza"])
        assert "POL-08" in polizas_base
        assert "POL-NEGATIVO" in polizas_base
        assert "POL-MORA-CERO" in polizas_base
        assert "POL-ANUAL-MEDIO" in polizas_base

        # Revisar Cancelación: Solo POL-99
        assert len(dfs["Revisar cancelación"]) == 1
        assert dfs["Revisar cancelación"].iloc[0]["Póliza"] == "POL-99"

        # Validación de Estado_Contacto ("No contactar" para 111111 en Alto Impacto y 999999 en Cancelación)
        assert dfs["Alto impacto"].loc[dfs["Alto impacto"]["Póliza"] == "POL-05", "Estado_Contacto"].iloc[0] == "No contactar"
        assert dfs["Revisar cancelación"].loc[dfs["Revisar cancelación"]["Póliza"] == "POL-99", "Estado_Contacto"].iloc[0] == "No contactar"

        # Validación de Asesor Principal en las hojas
        assert "Asesor principal" in dfs["Alto impacto"].columns
        assert "Asesor principal" in dfs["Base depurada"].columns
        assert dfs["Alto impacto"].loc[dfs["Alto impacto"]["Póliza"] == "POL-05", "Asesor principal"].iloc[0] == "Asesor_NG"
        assert dfs["Alto impacto"].loc[dfs["Alto impacto"]["Póliza"] == "POL-06", "Asesor principal"].iloc[0] == "Asesor_NB"

        # Validación de Resumen Comerciales y Hojas por Comercial
        df_res_com = result["resumen_comerciales"]
        assert not df_res_com.empty
        assert "Comercial" in df_res_com.columns
        assert "Total Cartera ($)" in df_res_com.columns
        assert "Foco de Acción Sugerido" in df_res_com.columns

        dfs_com = result["dfs_comerciales"]
        assert "AI - Asesor_NG" in dfs_com
        assert "AI - Asesor_NB" in dfs_com

        # Probar exportación a Excel con Resumen y Hojas por Asesor
        excel_bytes = export_cartera_to_excel(
            dfs,
            resumen_comerciales=df_res_com,
            dfs_comerciales=dfs_com,
        )
        assert len(excel_bytes) > 0

        excel_file = pd.ExcelFile(io.BytesIO(excel_bytes))
        sheet_names = set(excel_file.sheet_names)
        assert "Resumen Comerciales" in sheet_names
        assert "Base depurada" in sheet_names
        assert "Alto impacto" in sheet_names
        assert "Débitos fallidos" in sheet_names
        assert "Devoluciones" in sheet_names
        assert "Revisar cancelación" in sheet_names
        assert "AI - Asesor_NG" in sheet_names
        assert "AI - Asesor_NB" in sheet_names
        print("[PASS] test_pipeline_complete_flow passed")


if __name__ == "__main__":
    test_clean_id_and_numeric()
    test_pipeline_complete_flow()
    print("\nALL AUTOMATED TESTS PASSED SUCCESSFULLY!")
