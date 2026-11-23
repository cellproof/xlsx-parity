"""Rebuild fixtures/RT-features.xlsx. make test uses the committed workbook."""
import os

import xlsxwriter

HERE = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(HERE, "RT-features.xlsx")
wb = xlsxwriter.Workbook(path)
ws = wb.add_worksheet("Data")
bold = wb.add_format({"bold": True, "bg_color": "#FFFF00"})
date = wb.add_format({"num_format": "yyyy-mm-dd"})
ws.write("A1", "Header", bold)
ws.write("A2", 45000, date)
ws.write_formula("A3", "=A2+1", date, 45001)
ws.set_column("A:A", 30)
ws.merge_range("C1:D1", "merged", None)
ws.write_comment("A1", "a comment")
ws.data_validation("B2", {"validate": "list", "source": ["x", "y"]})
ws.freeze_panes(1, 0)
ws.write_url("B5", "https://example.com", string="link")
wb.define_name("MyName", "=Data!$A$2")
ws.write_formula("A4", "=MyName*2", None, 90000)
ws.conditional_format("A2:A3", {"type": "cell", "criteria": ">", "value": 0, "format": bold})
for i, value in enumerate([1, 3, 2]):
    ws.write(10 + i, 0, value)
chart = wb.add_chart({"type": "column"})
chart.add_series({"values": "=Data!$A$11:$A$13"})
ws.insert_chart("F2", chart)
wb.close()
print("wrote", path)
