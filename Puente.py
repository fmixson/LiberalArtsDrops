import pandas as pd

pd.set_option('display.max_columns', None)
# laDrops_df = pd.DataFrame(columns=headers)
file = 'Puente 100 103.csv'
Puente_df = pd.read_csv(file)
# Puente_df.index.name='Index'

print(Puente_df)

multiplesList = []
dropIndexList = []
for i in range(len(Puente_df)-1):
    if Puente_df.loc[i, 'Employee ID'] == Puente_df.loc[i+1, 'Employee ID']:
        Puente_df.drop(index=i)
        multiplesList.append(i)
        multiplesList.append(i+1)
print(multiplesList)
Puente_df.to_excel('Puente.xlsx')
PuenteClean_df = Puente_df.drop(multiplesList)
PuenteClean_df.to_excel('PuenteClean.xlsx')