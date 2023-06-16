import openpyxl
import requests as requests
from lxml import etree as ET

#функция для получения значения usd
def get_usd_value(sdate,svalue):
    sdate = sdate.strftime('%d.%m.%Y')
    response = requests.get(f"https://www.cbr.ru/currency_base/daily/?UniDbQuery.Posted=True&UniDbQuery.To={sdate}")
    if response.status_code == 200:
        table_start = response.text.find("Доллар США")
        usd_start = table_start + len("Доллар Сша") + len("</td>") + len("<td class=""> <td>")
        usd_value = float(response.text[usd_start:usd_start + 7].replace(",","."))
        result = str(round(svalue/usd_value,2))
        return result
    return "Server not working"


# Открытие файла Excel
workbook = openpyxl.load_workbook('test_input.xlsx', data_only=True)
sheet = workbook.active

# Создание корневого элемента CERTDATA
certdata = ET.Element('CERTDATA')

# Добавление элемента FILENAME
filename = ET.SubElement(certdata, 'FILENAME')
filename.text = sheet['B3'].value


# Создание элемента ENVELOPE
envelope = ET.SubElement(certdata, 'ENVELOPE')

# Чтение данных из таблицы Excel и создание элементов ECERT
for row in sheet.iter_rows(min_row=6, values_only=True):
    ecert = ET.SubElement(envelope, 'ECERT')
    certno = ET.SubElement(ecert, 'CERTNO')
    certno.text = str(row[0])
    certdate = ET.SubElement(ecert, 'CERTDATE')
    certdate.text = row[1].strftime('%Y-%m-%d')
    status = ET.SubElement(ecert, 'STATUS')
    status.text = row[2]
    iec = ET.SubElement(ecert, 'IEC')
    iec.text = str(row[3])
    expname = ET.SubElement(ecert, 'EXPNAME')
    expname.text = f'"{row[4]}"'
    billid = ET.SubElement(ecert, 'BILLID')
    billid.text = row[5]
    sdate = ET.SubElement(ecert, 'SDATE')
    sdate.text = row[6].strftime('%Y-%m-%d')
    scc = ET.SubElement(ecert, 'SCC')
    scc.text = row[7]
    svalue = ET.SubElement(ecert, 'SVALUE')
    svalue.text = str(row[8]).replace(',', '.')
    svalue_usd = ET.SubElement(ecert, 'SVALUEUSD')
    svalue_usd.text = get_usd_value(row[6],row[8])

# Создание объекта ElementTree и запись в файл
tree = ET.ElementTree(certdata)
tree.write('output_xml_v2.xml', encoding='utf-8', xml_declaration=True, pretty_print=True)
