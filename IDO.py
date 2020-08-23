from datetime import datetime

print('')
print('---------------------------------------')
print('Secciones IDO:')
print('(1) IDO completo')
print('(2) Generación')
print('(3) Intercambios internacionales')
print('(4) Disponibilidad')
print('(5) Sucesos')
print('(6) Costos')
print('(7) AGC programado')
print('(8) Hidrología')
print('(9) Eventos generadores')
print('(10) Despacho por Recurso/Unidad')
print('(11) Meteorología fuentes renovables')
print('---------------------------------------')
seccion = input('Elija la sección de IDO (0): ')
print('---------------------------------------')
print('Ingrese la fecha que desea descargar:')
day = input('Día (0): ')
month = input('Mes (0): ')
year = input('Año (0000): ')
print('---------------------------------------')

strDate = month + "/" + day + "/" + year
fecha_IDO = datetime.strptime(strDate, '%m/%d/%Y')
fecha_IDO = datetime.strftime(fecha_IDO, '%d/%m/%Y')

print('')
print('---------------------------------------')
print('Fecha del archivo IDO:',fecha_IDO)
print('---------------------------------------')

if seccion == '1':
    seccion = 'ido.aspx'
elif seccion == '2':
    seccion = 'generacion.aspx'
elif seccion == '3':
    seccion = 'intercambios.aspx'
elif seccion == '4':
    seccion = 'disponibilidad.aspx'
elif seccion == '5':
    print('Secciones Sucesos:')
    print('(1) Eventos de demanda no atendida')
    print('(2) Eventos de tensión fuera de rango')
    print('(3) Variación de frecuencia')
    print('(4) Desconexión automática de carga')
    print('---------------------------------------')
    seccion = input('Elija la sección de Sucesos (0): ')
    print('---------------------------------------')
    if seccion == '1':
        seccion = 'sucesos.aspx?q=demanda'
    elif seccion == '2':
        seccion = 'sucesos.aspx?q=tension'
    elif seccion == '3':
        seccion = 'sucesos.aspx?q=variaciones'
    elif seccion == '4':
        seccion = 'sucesos.aspx?q=desconexion'
elif seccion == '6':
    seccion = 'costos.aspx'
elif seccion == '7':
    seccion = 'programado.aspx'
elif seccion == '8':
    print('Secciones Hidrología:')
    print('(1) Reservas')
    print('(2) Vertimientos')
    print('(3) Aportes')
    print('(4) HSIN')
    print('---------------------------------------')
    seccion = input('Elija la sección de Hidrología (0): ')
    print('---------------------------------------')
    if seccion == '1':
        seccion = 'hidrologia.aspx?q=reservas'
    elif seccion == '2':
        seccion = 'hidrologia.aspx?q=vertimientos'
    elif seccion == '3':
        seccion = 'hidrologia.aspx?q=aportes'
        anexoname = year + '_' + month + '_' + day + '_' + 'AportesIDO.xls'
    elif seccion == '4':
        seccion = 'hidrologia.aspx?q=hsin'
elif seccion == '9':
    seccion = 'eventos.aspx'
elif seccion == '10':
    seccion = 'despacho.aspx'
elif seccion == '11':
    print('Secciones Meteorología FR:')
    print('(1) Plantas solares')
    print('(2) Plantas eólicas')
    print('---------------------------------------')
    seccion = input('Elija la sección de Meteorología FR (0): ')
    print('---------------------------------------')
    if seccion == '1':
        seccion = 'meteorologia.aspx?q=solares'
    elif seccion == '2':
        seccion = 'meteorologia.aspx?q=eolicas'

def descarga_ido(fecha_IDO,seccion):
    from selenium import webdriver
    import time
    driver = webdriver.Chrome()
    link = "http://ido.xm.com.co/ido/SitePages/" + seccion
    driver.get(link)
    if seccion == 'despacho.aspx':
        driver.find_element_by_id("filter-button").click()
        time.sleep(5)
        driver.find_element_by_id("ctl00_PlaceHolderPersonalizado_g_4fe9dfe7_92e4_414c_a152_3af85113b851_ctl00_boton_excel").click()
        time.sleep(2)
        driver.close()
    else:
        elem = driver.find_element_by_id("report-date")
        elem.send_keys(fecha_IDO)
        time.sleep(20)
        driver.find_element_by_id("filter-button").click()
        time.sleep(5)
        driver.find_element_by_id("ctl00_PlaceHolderPersonalizado_g_4fe9dfe7_92e4_414c_a152_3af85113b851_ctl00_boton_excel").click()
        time.sleep(2)
        driver.close()

descarga_ido(fecha_IDO,seccion)

import shutil
import os
import time

print("")
downloads = os.listdir('C:/Users/Juan Camilo/Downloads/')
file = 'C:/Users/Juan Camilo/Downloads/' + downloads[0]
print(file)
newnamefile = 'C:/Users/Juan Camilo/Downloads/' + 'Archivo_IDO.xls'
os.rename(file,newnamefile)
print('Archivo descargado en carpeta Descargas(Downloads)')
time.sleep(3)
