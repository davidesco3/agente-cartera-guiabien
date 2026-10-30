from pathlib import Path
import pandas as pd

EXAMPLES_DIR = Path(__file__).resolve().parent / "data" / "ejemplos"
EXAMPLES_DIR.mkdir(parents=True, exist_ok=True)

# 1. Cartera Actual
cartera_actual_data = [
    {
        "Nº ident.fis.1": "1017123456",
        "Póliza": "9800101",
        "Operación": "Cobro",
        "Medio Pago": "Débito",
        "Dias Cartera": 10,
        "Saldo a pagar": 185000,
        "Nombre Tomador": "Mateo Londoño Gómez",
        "Solución": "Autos Clásico",
        "Resp.Cobro": "",
    },
    {
        "Nº ident.fis.1": "1037654321",  # Este está en clientes_elite.csv predeterminado!
        "Póliza": "9800102",
        "Operación": "Cobro",
        "Medio Pago": "Transferencia",
        "Dias Cartera": 22,
        "Saldo a pagar": 890000,
        "Nombre Tomador": "Mariana Echeverri Villa",
        "Solución": "Salud Global",
        "Resp.Cobro": "",
    },
    {
        "Nº ident.fis.1": "71234567",   # En clientes_elite.csv
        "Póliza": "9800103",
        "Operación": "Cobro",
        "Medio Pago": "Efectivo",
        "Dias Cartera": 14,
        "Saldo a pagar": 450000,
        "Nombre Tomador": "Santiago Jaramillo Restrepo",
        "Solución": "Vida Individual",
        "Resp.Cobro": "",
    },
    {
        "Nº ident.fis.1": "1020304050",
        "Póliza": "9800104",
        "Operación": "Cobro",
        "Medio Pago": "Débito",
        "Dias Cartera": 28,
        "Saldo a pagar": 620000,
        "Nombre Tomador": "Camila Cardona Peña",
        "Solución": "Hogar Seguro",
        "Resp.Cobro": "Fondos insuficientes",
    },
    {
        "Nº ident.fis.1": "1030405060",
        "Póliza": "9800105",
        "Operación": "Cobro",
        "Medio Pago": "Débito",
        "Dias Cartera": 35,
        "Saldo a pagar": 340000,
        "Nombre Tomador": "Andrés Felipe Quintero",
        "Solución": "Mascotas",
        "Resp.Cobro": None,  # Vacío para probar fallback
    },
    {
        "Nº ident.fis.1": "900987654",
        "Póliza": "9800106",
        "Operación": "Cobro",
        "Medio Pago": "Transferencia",
        "Dias Cartera": 18,
        "Saldo a pagar": 4250000,  # > 3,000,000 Alto Impacto
        "Nombre Tomador": "Inversiones Medellín S.A.S.",
        "Solución": "Empresarial Multirriesgo",
        "Resp.Cobro": "",
    },
    {
        "Nº ident.fis.1": "43567890",
        "Póliza": "9800107",
        "Operación": "Cobro",
        "Medio Pago": "Tarjeta Crédito",
        "Dias Cartera": 52,  # >= 45 Alto Impacto
        "Saldo a pagar": 780000,
        "Nombre Tomador": "Gloria Inés Arango",
        "Solución": "Salud Tradicional",
        "Resp.Cobro": "",
    },
    {
        "Nº ident.fis.1": "12345678",  # En no_contactar_cache.csv
        "Póliza": "9800108",
        "Operación": "Cobro",
        "Medio Pago": "Transferencia",
        "Dias Cartera": 20,
        "Saldo a pagar": 510000,
        "Nombre Tomador": "Alejandro Botero Toro",
        "Solución": "Autos Total",
        "Resp.Cobro": "",
    },
    {
        "Nº ident.fis.1": "1015678912",
        "Póliza": "9800109",
        "Operación": "Devolución",  # Devolución!
        "Medio Pago": "Transferencia",
        "Dias Cartera": 5,
        "Saldo a pagar": -150000,
        "Nombre Tomador": "Valeria Osorio Vélez",
        "Solución": "Autos Básico",
        "Resp.Cobro": "",
    },
    {
        "Nº ident.fis.1": "1028765432",
        "Póliza": "9800110",
        "Operación": "Cobro",
        "Medio Pago": "Transferencia",
        "Dias Cartera": 19,
        "Saldo a pagar": 920000,
        "Nombre Tomador": "Sebastián Henao Ramos",
        "Solución": "Vida Grupo",
        "Resp.Cobro": "",
    },
]
df_actual = pd.DataFrame(cartera_actual_data)
df_actual.to_excel(EXAMPLES_DIR / "informe_cartera_actual_ejemplo.xlsx", index=False)

# 2. Cartera Semana Anterior
cartera_anterior_data = [
    {
        "Nº ident.fis.1": "1017123456",
        "Póliza": "9800101",
        "Nombre Tomador": "Mateo Londoño Gómez",
        "Saldo a pagar": 185000,
    },
    {
        "Nº ident.fis.1": "1037654321",
        "Póliza": "9800102",
        "Nombre Tomador": "Mariana Echeverri Villa",
        "Saldo a pagar": 890000,
    },
    {
        "Nº ident.fis.1": "1099887766",
        "Póliza": "9800999",  # NO está en actual, pero SÍ está en recaudos (pagó)
        "Nombre Tomador": "Patricia Correa Gil",
        "Saldo a pagar": 650000,
    },
    {
        "Nº ident.fis.1": "98765432",  # NO está en actual y NO está en recaudos -> Posible Cancelación! (Y está en no_contactar)
        "Póliza": "9800888",
        "Nombre Tomador": "Construcciones Antioquia S.A.S.",
        "Saldo a pagar": 1200000,
    },
]
df_anterior = pd.DataFrame(cartera_anterior_data)
df_anterior.to_excel(EXAMPLES_DIR / "informe_cartera_semana_anterior_ejemplo.xlsx", index=False)

# 3. Recaudos
recaudos_data = [
    {
        "Nº ident.fis.1": "1099887766",
        "Póliza": "9800999",
        "Valor Recaudo": 650000,
        "Fecha Recaudo": "2026-10-25",
    }
]
df_recaudos = pd.DataFrame(recaudos_data)
df_recaudos.to_excel(EXAMPLES_DIR / "informe_recaudos_ejemplo.xlsx", index=False)

print("Sample Excel files generated successfully in:", EXAMPLES_DIR)
