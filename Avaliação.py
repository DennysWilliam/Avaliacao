import pandas as pd

listaSites = pd.read_excel(r'C:\Users\Dennys\Desktop\Avaliação\SiteList.xlsx')
resultados = pd.read_excel(r'C:\Users\Dennys\Desktop\Avaliação\Results.xlsx')

listaSites_2023 = listaSites[listaSites['Year'] == 2023]
resultados_2023 = resultados[resultados['Year'] == 2023]
relatorio = pd.DataFrame(listaSites_2023, columns=['Site Name', 'State'])
relatorio = pd.DataFrame(resultados_2023, columns=['Site ID', 'Equipment', 'Signal (%)', 'Quality (0-10)', 'Mbps'])

relatorio = pd.merge(listaSites_2023, resultados_2023, on='Site ID', how='inner')
relatorio = relatorio.sort_values(by='Site ID')
print(relatorio)

print ('Sites com alerta ativos:')
print (resultados_2023[resultados_2023['Alerts'] == 'Sim'])
print ('\nSites com 0 de qualidade:')
print (resultados_2023[resultados_2023['Quality (0-10)'] == 0])
print ('\nSites com mais de 80 Mbps:')
print (resultados_2023[resultados_2023['Mbps'] > 80])
print ('\nSites com menos de 10 Mbps:')
print (resultados_2023[resultados_2023['Mbps'] < 10])

sitesSemResultado = listaSites_2023[listaSites_2023['Site ID'].isin(resultados_2023['Site ID']).to_list()]
print ("\nSites que não estão presentes no Results: ")
print (sitesSemResultado["Site ID"])

with pd.ExcelWriter('Relatorio.xlsx') as writer:
    relatorio.to_excel(writer, index=False)