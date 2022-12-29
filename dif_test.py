import pandas as pd
from datetime import datetime

df_users = pd.read_excel('user_admin.xlsx')
df_users = df_users.fillna('')
# print(df_users.head())

# df_scans = pd.read_excel('Данные по пользователям и сканам 2022.xlsx')
# df_scans = df_scans.fillna('')
# print(df_scans.head())

exclude_users = ['kazah89', 'sanin, ''samoilov', 'axorindustry', 'kreknina', 'zeykin', 'berdnikova', 'ostashenko',
                 'skalar', 'test', 'malyigor', 'ihormaly', 'axor']
exclude_list = set()

for email in df_users['E-Mail']:
    for i in exclude_users:
        if i in email:
            exclude_list.add(email)

print(exclude_list)
