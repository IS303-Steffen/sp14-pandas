'''
OPTIONAL AI GUIDANCE PROMPT
---------------------------
I am a student in an introductory Python class. I am learning many coding
principles for the very first time. I am going to paste in the instructions
to a practice problem that my professor gave me to try before class.
Please be my kind tutor and walk me through how to solve the problem step
by step.

Don't just give me the full solution all at once (unless I later ask for
it). Instead, help me work through it gradually, with clear explanations
and small, easy-to-understand examples. Please use everyday language and
explain things in a simple, friendly way.

INSTRUCTIONS:
-------------
Filter sales_q1.csv for rows where Month equals "Mar" using .loc. Append these
rows to the sheet Q1_Summary in summary.xlsx. Save as a new workbook called
summary_updated.xlsx.
'''

# Here is one potential solution. Remember there are often many different
# ways to solve a problem, so your solution may not look exactly the same.

import pandas as pd
from openpyxl import load_workbook

df = pd.read_csv('sales_q1.csv')
mar_rows = df.loc[df['Month'] == 'Mar']

wb = load_workbook('summary.xlsx')
ws = wb['Q1_Summary']
for _, row in mar_rows.iterrows():
    ws.append([row['Month'], row['Region'], row['Units'], row['Revenue']])
wb.save('summary_updated.xlsx')
