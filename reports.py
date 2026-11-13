import html
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from PySide6.QtGui import QTextDocument, QPageSize, QPageLayout
from PySide6.QtCore import QMarginsF
from PySide6.QtPrintSupport import QPrinter

from calculations import get_today_jalali_str


def export_to_excel(file_path, report_title, headers, rows, warehouse_name="انبار مرکزی"):
    """ساخت فایل اکسل راست‌چین با فونت B Titr برای عنوان‌ها و B Nazanin برای متن"""
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "گزارش انبار"
        ws.views.sheetView[0].rightToLeft = True

        font_title = Font(name="B Titr", size=14, bold=True, color="FFFFFF")
        font_sub = Font(name="B Nazanin", size=12, bold=True, color="1E272E")
        font_header = Font(name="B Titr", size=11, bold=True, color="FFFFFF")
        font_cell = Font(name="B Nazanin", size=12, bold=False, color="000000")

        fill_title = PatternFill("solid", fgColor="0FB9B1")
        fill_sub = PatternFill("solid", fgColor="E8F8F5")
        fill_header = PatternFill("solid", fgColor="2F3542")
        fill_zebra = PatternFill("solid", fgColor="F8F9FA")
        fill_white = PatternFill("solid", fgColor="FFFFFF")

        thin_side = Side(border_style="thin", color="CED6E0")
        border_box = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

        col_count = max(1, len(headers))
        last_col_letter = get_column_letter(col_count)

        ws.merge_cells(f"A1:{last_col_letter}1")
        ws["A1"] = f"{warehouse_name}  |  {report_title}"
        ws["A1"].font = font_title
        ws["A1"].fill = fill_title
        ws["A1"].alignment = align_center
        ws.row_dimensions[1].height = 34

        ws.merge_cells(f"A2:{last_col_letter}2")
        ws["A2"] = f"تاریخ تهیه گزارش: {get_today_jalali_str()}   |   تعداد کل ردیف‌ها: {len(rows)} مورد"
        ws["A2"].font = font_sub
        ws["A2"].fill = fill_sub
        ws["A2"].alignment = align_center
        ws.row_dimensions[2].height = 26

        header_row_idx = 4
        for col_idx, h_text in enumerate(headers, start=1):
            cell = ws.cell(row=header_row_idx, column=col_idx, value=str(h_text))
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = border_box
        ws.row_dimensions[header_row_idx].height = 30

        for r_offset, row_data in enumerate(rows):
            r_num = header_row_idx + 1 + r_offset
            row_fill = fill_zebra if r_offset % 2 == 1 else fill_white
            for c_idx, val in enumerate(row_data, start=1):
                cell = ws.cell(row=r_num, column=col_idx if False else c_idx, value=str(val))
                cell.font = font_cell
                cell.fill = row_fill
                cell.alignment = align_center
                cell.border = border_box
            ws.row_dimensions[r_num].height = 25

        for col_idx in range(1, col_count + 1):
            col_letter = get_column_letter(col_idx)
            max_len = len(str(headers[col_idx - 1]))
            for row_data in rows:
                if col_idx - 1 < len(row_data):
                    val_len = len(str(row_data[col_idx - 1]))
                    if val_len > max_len:
                        max_len = val_len
            ws.column_dimensions[col_letter].width = max(15, min(max_len + 7, 44))

        wb.save(file_path)
        return True, f"فایل اکسل با موفقیت در مسیر زیر ذخیره شد:\n{file_path}"
    except PermissionError:
        return False, "فایل اکسل قبلی در برنامه Excel باز است! لطفاً ابتدا آن را ببندید و دوباره تلاش کنید."
    except Exception as e:
        return False, f"خطا در ساخت فایل اکسل:\n{str(e)}"


def export_to_pdf(file_path, report_title, headers, rows, warehouse_name="انبار مرکزی"):
    """ساخت فایل PDF رسمی، فارسی و راست‌چین با فونت B Titr و B Nazanin"""
    try:
        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(file_path)
        printer.setPageSize(QPageSize(QPageSize.A4))
        printer.setPageOrientation(QPageLayout.Landscape)
        printer.setPageMargins(QMarginsF(10, 10, 10, 10), QPageLayout.Millimeter)

        th_html = "".join(f"<th>{html.escape(str(h))}</th>" for h in headers)

        body_rows_html = []
        for idx, r in enumerate(rows):
            bg_color = "#f8f9fa" if idx % 2 == 1 else "#ffffff"
            tds = "".join(f"<td>{html.escape(str(cell))}</td>" for cell in r)
            body_rows_html.append(f"<tr style='background-color: {bg_color};'>{tds}</tr>")

        if not body_rows_html:
            body_rows_html.append(
                f"<tr><td colspan='{len(headers)}'>هیچ رکوردی برای نمایش وجود ندارد.</td></tr>"
            )

        html_content = f"""
        <!DOCTYPE html>
        <html lang="fa" dir="rtl">
        <head>
            <meta charset="utf-8">
            <style>
                body {{
                    font-family: 'B Nazanin', Tahoma, sans-serif;
                    direction: rtl;
                    text-align: right;
                    color: #1e272e;
                }}
                .header-box {{
                    background-color: #1e272e;
                    color: #ffffff;
                    padding: 12px;
                    text-align: center;
                    border-bottom: 3px solid #0fb9b1;
                    margin-bottom: 12px;
                }}
                .title {{
                    font-family: 'B Titr', Tahoma, sans-serif;
                    font-size: 15pt;
                    font-weight: bold;
                    margin-bottom: 5px;
                }}
                .meta {{
                    font-family: 'B Nazanin', Tahoma, sans-serif;
                    font-size: 11pt;
                    color: #d2dae2;
                }}
                table {{
                    width: 100%;
                    border-collapse: collapse;
                    font-family: 'B Nazanin', Tahoma, sans-serif;
                    font-size: 11pt;
                }}
                th {{
                    font-family: 'B Titr', Tahoma, sans-serif;
                    font-size: 10pt;
                    background-color: #2f3542;
                    color: #ffffff;
                    font-weight: bold;
                    padding: 8px 5px;
                    border: 1px solid #7f8c8d;
                    text-align: center;
                }}
                td {{
                    font-family: 'B Nazanin', Tahoma, sans-serif;
                    font-size: 11pt;
                    padding: 6px 5px;
                    border: 1px solid #bdc3c7;
                    text-align: center;
                }}
            </style>
        </head>
        <body>
            <div class="header-box">
                <div class="title">{html.escape(warehouse_name)} - {html.escape(report_title)}</div>
                <div class="meta">تاریخ گزارش: {html.escape(get_today_jalali_str())} &nbsp;|&nbsp; تعداد رکوردها: {len(rows)} مورد</div>
            </div>
            <table>
                <thead>
                    <tr>{th_html}</tr>
                </thead>
                <tbody>
                    {"".join(body_rows_html)}
                </tbody>
            </table>
        </body>
        </html>
        """

        doc = QTextDocument()
        doc.setHtml(html_content)
        doc.print_(printer)
        return True, f"فایل PDF با موفقیت در مسیر زیر ذخیره شد:\n{file_path}"
    except Exception as e:
        return False, f"خطا در ساخت فایل PDF:\n{str(e)}"