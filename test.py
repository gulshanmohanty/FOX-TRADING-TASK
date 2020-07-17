import xlwt
import xlrd
import requests
import time

wb = xlwt.Workbook()
ws = wb.add_sheet("weather")
ws.write(0,0,"city name")
ws.write(0,1,"temprature in kelvin ")
ws.write(0,2,"temp unit")
ws.write(0,3,"humidity")
ws.write(0,4,"weather")
inputWorkbook = xlrd.open_workbook("test.xls")#it is used to select test.xml
inputWorksheet = inputWorkbook.sheet_by_index(1)#it is used to select sheet by index
rows=inputWorksheet.nrows#it is used to count rows

api_address = 'http://api.openweathermap.org/data/2.5/weather?appid=255108f14faae1200038496340eb9e59&q='

#var=1
#while var ==1:#(to run the loop continously)
i=1
for r in range (0,rows):
    
    inputWorkbook = xlrd.open_workbook("test.xls")
    inputWorksheet = inputWorkbook.sheet_by_index(1)
    print(inputWorksheet.cell_value(i,0))

        
    city = inputWorksheet.cell_value(i,0)#it is used to take city names from test.xml
    url = api_address + city
    data = requests.get(url).json()
    #JavaScript Object Notation (JSON) is a standard text-based format for representing structured data.
    #temprature_data= data['main']['temp']

    #print(data)#(optional) it is used to read all data about weather.
    #print(temprature_data)#(optional) it is used to check only temprature of city.
    
    i=i+1
    ws.write(i,0,city)
    ws.write(i,1,data['main']['temp'])#c= k-273.15 , f = k-459.67
    ws.write(i,3,data['main']['humidity'])
    ws.write(i,2,'kelvin')
    ws.write(i,4,data['weather'][0]['main'])

    
    wb.save("weather.xls")#it is used to save the xml file

    time.sleep(1)
    if i==rows:
        break
print('weather update complete')

