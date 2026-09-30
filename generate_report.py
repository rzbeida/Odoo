import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Load all data files
print("جاري تحميل البيانات...")
try:
    invoices_df = pd.read_excel('Pivot Invoices Analysis (account.invoice.report).xlsx')
    purchase_df = pd.read_excel('Purchase Order (purchase.order).xlsx')
    sales_df = pd.read_excel('Sales Order (sale.order).xlsx')
    team_member_df = pd.read_excel('Sales Team Member (crm.team.member).xlsx')
    profit_loss_df = pd.read_excel('profit_and_loss.xlsx')
    print("✓ تم تحميل جميع الملفات بنجاح")
except Exception as e:
    print(f"خطأ في تحميل الملفات: {e}")
    exit()

# Create a new workbook
wb = openpyxl.Workbook()
wb.remove(wb.active)

# Define styles
title_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
title_font = Font(name='Arial', size=14, bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
header_font = Font(name='Arial', size=11, bold=True, color="FFFFFF")
border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

def style_sheet(ws, title):
    """Add title and styling to sheet"""
    ws.insert_rows(1)
    ws['A1'] = title
    ws['A1'].font = title_font
    ws['A1'].fill = title_fill
    ws.merge_cells('A1:Z1')
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 25
    
    # Style headers
    for cell in ws[2]:
        if cell.value:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border

# ============ Sheet 1: ملخص تنفيذي ============
ws_summary = wb.create_sheet("ملخص تنفيذي", 0)

summary_data = [
    ["تقرير شهر سبتمبر 2024", ""],
    ["", ""],
    ["البيان", "القيمة"],
    ["إجمالي الفواتير", f"${invoices_df['Total'].sum() if 'Total' in invoices_df.columns else 0:,.2f}"],
    ["إجمالي أوامر الشراء", f"${purchase_df.iloc[:, 2].sum() if len(purchase_df.columns) > 2 else 0:,.2f}"],
    ["إجمالي أوامر البيع", f"${sales_df.iloc[:, 2].sum() if len(sales_df.columns) > 2 else 0:,.2f}"],
    ["عدد أعضاء فريق المبيعات", len(team_member_df)],
]

for row in summary_data:
    ws_summary.append(row)

ws_summary['A1'].font = Font(name='Arial', size=16, bold=True, color="1F4E78")
ws_summary.column_dimensions['A'].width = 30
ws_summary.column_dimensions['B'].width = 25

# ============ Sheet 2: تحليل الفواتير ============
ws_invoices = wb.create_sheet("تحليل الفواتير")
for r_idx, row in enumerate(dataframe_to_rows(invoices_df, index=False, header=True), 1):
    for c_idx, value in enumerate(row, 1):
        ws_invoices.cell(row=r_idx, column=c_idx, value=value)
style_sheet(ws_invoices, "تحليل الفواتير (Invoices)")
for col in ws_invoices.columns:
    ws_invoices.column_dimensions[col[0].column_letter].width = 18

# ============ Sheet 3: أوامر الشراء ============
ws_purchase = wb.create_sheet("أوامر الشراء")
for r_idx, row in enumerate(dataframe_to_rows(purchase_df, index=False, header=True), 1):
    for c_idx, value in enumerate(row, 1):
        ws_purchase.cell(row=r_idx, column=c_idx, value=value)
style_sheet(ws_purchase, "أوامر الشراء (Purchase Orders)")
for col in ws_purchase.columns:
    ws_purchase.column_dimensions[col[0].column_letter].width = 18

# ============ Sheet 4: أوامر البيع ============
ws_sales = wb.create_sheet("أوامر البيع")
for r_idx, row in enumerate(dataframe_to_rows(sales_df, index=False, header=True), 1):
    for c_idx, value in enumerate(row, 1):
        ws_sales.cell(row=r_idx, column=c_idx, value=value)
style_sheet(ws_sales, "أوامر البيع (Sales Orders)")
for col in ws_sales.columns:
    ws_sales.column_dimensions[col[0].column_letter].width = 18

# ============ Sheet 5: أعضاء فريق المبيعات ============
ws_team = wb.create_sheet("فريق المبيعات")
for r_idx, row in enumerate(dataframe_to_rows(team_member_df, index=False, header=True), 1):
    for c_idx, value in enumerate(row, 1):
        ws_team.cell(row=r_idx, column=c_idx, value=value)
style_sheet(ws_team, "أعضاء فريق المبيعات (Sales Team Members)")
for col in ws_team.columns:
    ws_team.column_dimensions[col[0].column_letter].width = 18

# ============ Sheet 6: تحليل الربح والخسارة ============
ws_pl = wb.create_sheet("الربح والخسارة")
for r_idx, row in enumerate(dataframe_to_rows(profit_loss_df, index=False, header=True), 1):
    for c_idx, value in enumerate(row, 1):
        ws_pl.cell(row=r_idx, column=c_idx, value=value)
style_sheet(ws_pl, "تحليل الربح والخسارة (Profit & Loss)")
for col in ws_pl.columns:
    ws_pl.column_dimensions[col[0].column_letter].width = 18

# ============ Sheet 7: أكثر مندوبي التحصيل ============
ws_top_collectors = wb.create_sheet("أكثر مندوبي التحصيل")
top_collectors = team_member_df.nlargest(10, team_member_df.columns[1] if len(team_member_df.columns) > 1 else team_member_df.columns[0])
for r_idx, row in enumerate(dataframe_to_rows(top_collectors, index=False, header=True), 1):
    for c_idx, value in enumerate(row, 1):
        ws_top_collectors.cell(row=r_idx, column=c_idx, value=value)
style_sheet(ws_top_collectors, "أكثر مندوبي التحصيل أداءً")
for col in ws_top_collectors.columns:
    ws_top_collectors.column_dimensions[col[0].column_letter].width = 18

# ============ Sheet 8: أكثر الإداريين مبيعاً ============
ws_top_sales = wb.create_sheet("أكثر الإداريين مبيعاً")
top_sales = sales_df.nlargest(10, sales_df.columns[2] if len(sales_df.columns) > 2 else sales_df.columns[0])
for r_idx, row in enumerate(dataframe_to_rows(top_sales, index=False, header=True), 1):
    for c_idx, value in enumerate(row, 1):
        ws_top_sales.cell(row=r_idx, column=c_idx, value=value)
style_sheet(ws_top_sales, "أكثر الإداريين مبيعاً")
for col in ws_top_sales.columns:
    ws_top_sales.column_dimensions[col[0].column_letter].width = 18

# Save the report
output_file = f"تقرير_شهر_سبتمبر_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
wb.save(output_file)
print(f"✓ تم إنشاء التقرير بنجاح: {output_file}")
