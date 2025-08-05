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
Read customers.csv. Use the dictionary provided in the code comment
to map the state_code column to full state names. Store the result in
a new column called state_name, then print the updated DataFrame.
'''

# Here is one potential solution. Remember there are often many different
# ways to solve a problem, so your solution may not look exactly the same.

# Dictionary for mapping state codes to names
state_map = {
    'CA': 'California',
    'TX': 'Texas',
    'CO': 'Colorado',
    'NY': 'New York'
}
import pandas as pd

state_map = {
    'CA': 'California',
    'TX': 'Texas',
    'CO': 'Colorado',
    'NY': 'New York'
}

df = pd.read_csv('customers.csv')
df['state_name'] = df['state_code'].map(state_map)
print(df)
