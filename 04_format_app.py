from openpyxl import load_workbook
from openpyxl.styles import Font

wb = load_workbook("reports/report.xlsx")
sh = wb["Report"]

sh["A1"] = "Reporte de Hours de IG"
sh["A2"] = "Hours Internas y Subcontratadas"

sh["A1"].font = Font("Segoe UI", bold=True, size=12)
sh["A2"].font = Font("Segoe UI", bold=True, size=8)

wb.save("reports/report_Hours.xlsx")
