import pandas as pd

exel = pd.read_excel('Kopia_Kopia_NOVYE_Vygruzka_SPO_GPOU_Kuznetskiy_industrialny_tekhnikum_1546sht.xlsx', engine='openpyxl')
exel = exel.head(5)

