# Automate with Python – Full Course for Beginners
# https://youtu.be/PXMJ6FS7llk?si=bJs8LTbxuvonaIch

import pandas as pd

df = pd.read_excel("db/Table_DB_240210.xlsx")

df = df[["Specialty Code", "Hour Type", "Hours"]]
# print(df)

pivot_table = df.pivot_table(
    index="Hour Type", columns="Specialty Code", values="Hours", aggfunc="sum"
)

# Workbook name, tab name, start row (row where the pivot is going to start)
pivot_table.to_excel("reports/pivot_table.xlsx", "Report", startrow=4)
