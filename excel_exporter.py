import io
from typing import Dict, Optional
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

from config import (
    BRAND_PRIMARY,
    EXCEL_ALERT_BG,
    EXCEL_ALERT_FG,
)


def export_cartera_to_excel(
    dfs: Dict[str, pd.DataFrame],
    resumen_comerciales: Optional[pd.DataFrame] = None,
    dfs_comerciales: Optional[Dict[str, pd.DataFrame]] = None,
) -> bytes:
    """
    Exporta el informe completo de cartera a un archivo Excel profesional utilizando XlsxWriter.
    Incluye:
    1. Hoja ejecutiva 'Resumen Comerciales' con totales y orientación a la acción.
    2. Las 5 hojas estándar depuradas (Base depurada, Alto impacto, Débitos fallidos, Devoluciones, Revisar cancelación).
    3. Hojas individuales por cada comercial con casos de Alto Impacto ('AI - {Comercial}').
    4. Formato condicional y nativo de resaltado rojo claro (#FFC7CE) en filas de 'No contactar'.
    """
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        # Estilos y formatos institucionales
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

        header_resumen_format = workbook.add_format({
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": "#0F172A",
            "border": 1,
            "border_color": "#94A3B8",
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

        cell_currency = workbook.add_format({
            "font_name": "Segoe UI",
            "font_size": 10,
            "valign": "vcenter",
            "border": 1,
            "border_color": "#E5E7EB",
            "num_format": "$ #,##0",
            "align": "right",
        })

        cell_integer = workbook.add_format({
            "font_name": "Segoe UI",
            "font_size": 10,
            "valign": "vcenter",
            "border": 1,
            "border_color": "#E5E7EB",
            "num_format": "#,##0",
            "align": "right",
        })

        cell_pct = workbook.add_format({
            "font_name": "Segoe UI",
            "font_size": 10,
            "valign": "vcenter",
            "border": 1,
            "border_color": "#E5E7EB",
            "num_format": "0.0%",
            "align": "right",
        })

        total_row_currency = workbook.add_format({
            "bold": True,
            "font_name": "Segoe UI",
            "font_size": 10,
            "bg_color": "#F1F5F9",
            "valign": "vcenter",
            "border": 1,
            "border_color": "#94A3B8",
            "num_format": "$ #,##0",
            "align": "right",
        })

        total_row_label = workbook.add_format({
            "bold": True,
            "font_name": "Segoe UI",
            "font_size": 10,
            "bg_color": "#F1F5F9",
            "valign": "vcenter",
            "border": 1,
            "border_color": "#94A3B8",
            "align": "left",
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

        # =====================================================================
        # 1. HOJA PRINCIPAL: RESUMEN EJECUTIVO POR COMERCIAL
        # =====================================================================
        if resumen_comerciales is not None and not resumen_comerciales.empty:
            sheet_name = "Resumen Comerciales"
            ws_resumen = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = ws_resumen
            ws_resumen.freeze_panes(1, 0)

            cols_resumen = list(resumen_comerciales.columns)
            for c_idx, c_name in enumerate(cols_resumen):
                ws_resumen.write(0, c_idx, c_name, header_resumen_format)

            for r_idx, row in enumerate(resumen_comerciales.itertuples(index=False), start=1):
                for c_idx, (c_name, val) in enumerate(zip(cols_resumen, row)):
                    if "($)" in c_name:
                        ws_resumen.write(r_idx, c_idx, float(val) if pd.notna(val) else 0.0, cell_currency)
                    elif "%" in c_name:
                        ws_resumen.write(r_idx, c_idx, (float(val) / 100.0) if pd.notna(val) else 0.0, cell_pct)
                    elif c_name in ("Pólizas en Mora", "Casos Alto Impacto", "Débitos Fallidos"):
                        ws_resumen.write(r_idx, c_idx, int(val) if pd.notna(val) else 0, cell_integer)
                    else:
                        ws_resumen.write(r_idx, c_idx, str(val) if pd.notna(val) else "", cell_format)

            # Fila de Totales
            tot_row = len(resumen_comerciales) + 1
            ws_resumen.write(tot_row, 0, "TOTAL CARTERA", total_row_label)
            if "Total Cartera ($)" in cols_resumen:
                idx_tc = cols_resumen.index("Total Cartera ($)")
                sum_tc = resumen_comerciales["Total Cartera ($)"].sum()
                ws_resumen.write(tot_row, idx_tc, sum_tc, total_row_currency)
            if "Casos Alto Impacto" in cols_resumen:
                idx_ai = cols_resumen.index("Casos Alto Impacto")
                sum_ai = resumen_comerciales["Casos Alto Impacto"].sum()
                ws_resumen.write(tot_row, idx_ai, int(sum_ai), total_row_currency)
            if "Saldo Alto Impacto ($)" in cols_resumen:
                idx_sai = cols_resumen.index("Saldo Alto Impacto ($)")
                sum_sai = resumen_comerciales["Saldo Alto Impacto ($)"].sum()
                ws_resumen.write(tot_row, idx_sai, sum_sai, total_row_currency)
            if "% Cartera" in cols_resumen:
                idx_pct = cols_resumen.index("% Cartera")
                ws_resumen.write(tot_row, idx_pct, 1.0, cell_pct)

            # Ajuste de columnas en el resumen
            for c_idx, c_name in enumerate(cols_resumen):
                max_len = max(len(str(c_name)), 14)
                if c_name == "Foco de Acción Sugerido":
                    max_len = 50
                ws_resumen.set_column(c_idx, c_idx, max_len + 3)

        # =====================================================================
        # 2. HOJAS ESTÁNDAR Y HOJAS INDIVIDUALES POR COMERCIAL
        # =====================================================================
        all_sheets_to_write = dict(dfs)
        if dfs_comerciales:
            for s_name, df_c in dfs_comerciales.items():
                all_sheets_to_write[s_name] = df_c

        for sheet_name, df in all_sheets_to_write.items():
            worksheet = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = worksheet
            worksheet.freeze_panes(1, 0)

            columns = list(df.columns)
            num_cols = len(columns)
            num_rows = len(df)

            # Escribir encabezados con formato corporativo
            for col_idx, col_name in enumerate(columns):
                worksheet.write(0, col_idx, col_name, header_format)

            if num_cols == 0:
                continue

            nc_col_idx = columns.index("Estado_Contacto") if "Estado_Contacto" in columns else -1

            # Escribir datos celda por celda
            for row_offset, row_tuple in enumerate(df.itertuples(index=False), start=1):
                is_no_contactar = (nc_col_idx != -1 and str(row_tuple[nc_col_idx]).strip() == "No contactar")
                row_style = no_contactar_format if is_no_contactar else cell_format

                for col_idx, val in enumerate(row_tuple):
                    if pd.isna(val) or val is None:
                        val_str = ""
                    else:
                        val_str = val
                    worksheet.write(row_offset, col_idx, val_str, row_style)

            # Ajustar anchos de columnas automáticamente
            for col_idx, col_name in enumerate(columns):
                max_len = len(str(col_name))
                if num_rows > 0:
                    sample_series = df[col_name].dropna().astype(str)
                    if not sample_series.empty:
                        max_data_len = sample_series.map(len).max()
                        max_len = max(max_len, max_data_len)
                col_width = min(max(max_len + 3, 12), 48)
                worksheet.set_column(col_idx, col_idx, col_width)

            # Regla Condicional Dinámica en Excel para persistencia ante ediciones
            if num_rows > 0 and nc_col_idx != -1:
                col_letter = xl_col_to_name(nc_col_idx)
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
