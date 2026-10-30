import io
from typing import Dict
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

from config import (
    BRAND_PRIMARY,
    EXCEL_ALERT_BG,
    EXCEL_ALERT_FG,
)


def export_cartera_to_excel(dfs: Dict[str, pd.DataFrame]) -> bytes:
    """
    Exporta los 5 DataFrames depurados a un archivo Excel en memoria utilizando XlsxWriter.
    Aplica formato institucional GuíaBien en encabezados y regla condicional de resaltado
    rojo claro (#FFC7CE) con texto rojo oscuro (#9C0006) en toda fila donde
    'Estado_Contacto' sea 'No contactar'.
    """
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        # Estilos y formatos
        header_format = workbook.add_format({
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": BRAND_PRIMARY,
            "border": 1,
            "border_color": "#D1D5DB",
            "align": "center",
            "valign": "vcenter",
            "font_name": "Segoe UI",
            "font_size": 11,
        })

        cell_format = workbook.add_format({
            "font_name": "Segoe UI",
            "font_size": 10,
            "valign": "vcenter",
            "border": 1,
            "border_color": "#E5E7EB",
        })

        no_contactar_format = workbook.add_format({
            "bg_color": EXCEL_ALERT_BG,
            "font_color": EXCEL_ALERT_FG,
            "font_name": "Segoe UI",
            "font_size": 10,
            "valign": "vcenter",
            "border": 1,
            "border_color": "#FCA5A5",
        })

        # Iterar sobre cada hoja requerida
        for sheet_name, df in dfs.items():
            # Crear la hoja en el libro
            worksheet = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = worksheet

            # Inmovilizar fila superior (encabezados)
            worksheet.freeze_panes(1, 0)

            columns = list(df.columns)
            num_cols = len(columns)
            num_rows = len(df)

            # Escribir encabezados con formato corporativo GuíaBien
            for col_idx, col_name in enumerate(columns):
                worksheet.write(0, col_idx, col_name, header_format)

            if num_cols == 0:
                continue

            # Localizar columna Estado_Contacto
            nc_col_idx = columns.index("Estado_Contacto") if "Estado_Contacto" in columns else -1

            # Escribir datos celda por celda aplicando estilo nativo
            for row_offset, row_tuple in enumerate(df.itertuples(index=False), start=1):
                is_no_contactar = (nc_col_idx != -1 and str(row_tuple[nc_col_idx]).strip() == "No contactar")
                row_style = no_contactar_format if is_no_contactar else cell_format

                for col_idx, val in enumerate(row_tuple):
                    if pd.isna(val) or val is None:
                        val_str = ""
                    else:
                        val_str = val
                    worksheet.write(row_offset, col_idx, val_str, row_style)

            # Ajustar anchos de columnas automáticamente según longitud de texto
            for col_idx, col_name in enumerate(columns):
                max_len = len(str(col_name))
                if num_rows > 0:
                    sample_series = df[col_name].dropna().astype(str)
                    if not sample_series.empty:
                        max_data_len = sample_series.map(len).max()
                        max_len = max(max_len, max_data_len)

                col_width = min(max(max_len + 3, 12), 48)
                worksheet.set_column(col_idx, col_idx, col_width)

            # Aplicar además Regla Condicional Dinámica en Excel para persistencia ante ediciones
            if num_rows > 0 and nc_col_idx != -1:
                col_letter = xl_col_to_name(nc_col_idx)

                # Rango de datos: fila 1 (segunda fila en Excel) hasta num_rows (última fila de datos)
                first_row = 1
                last_row = num_rows
                first_col = 0
                last_col = num_cols - 1

                formula = f'=${col_letter}2="No contactar"'

                worksheet.conditional_format(
                    first_row,
                    first_col,
                    last_row,
                    last_col,
                    {
                        "type": "formula",
                        "criteria": formula,
                        "format": no_contactar_format,
                    },
                )

    output.seek(0)
    return output.getvalue()
