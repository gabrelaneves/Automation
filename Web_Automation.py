from selenium import webdriver
from selenium.webdriver.common.keys import Keys

# -------------------- SELENIUM - WEB AUTOMATION ----------------------------------
#--------------------------------------------------------------
navegador = webdriver.Chrome()
#Golden
navegador.get('https://www.melhorcambio.com/ouro-hoje')
golden_price = navegador.find_element_by_xpath('//*[@id="comercial"]').get_attribute('value')
golden_price = golden_price.replace(",",".")
print('%.2f' % float(golden_price))
#Dolar
navegador.get('https://www.google.com/search?q=cota%C3%A7%C3%A3o+dolar')
dolar_price = navegador.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print( '%.2f' % float(dolar_price))
#Euro
navegador.get('https://www.google.com/search?q=cota%C3%A7%C3%A3o+euro')
euro_price = navegador.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
print('%.2f' % float(euro_price))

#--------------------------------------------------------------
import pandas as pd
df = pd.read_excel("Produtos.xlsx")
df.loc[df["Moeda"] == "Dólar", "Cotação"] = float(dolar_price)
df.loc[df["Moeda"] == "Euro", "Cotação"] = float(euro_price)
df.loc[df["Moeda"] == "Ouro", "Cotação"] = float(golden_price)
df["Preço Final"]= pd.to_numeric(df["Preço Final"],errors ="coerce")
df["Preço Base Reais"] = df["Cotação"]* df["Preço Base Original"]
df["Preço Final"] = df["Preço Base Reais"]*df["Ajuste"]
df["Preço Final"] = df["Preço Final"].map("{:.2f}".format)


df.info()

#Exportando para o Excel
df.to_excel("Produtos atu.xlsx", index = False)

#--------------------------------------------------------------
