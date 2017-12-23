dim myObj
Dim myClass
Set myObj = CreateObject("BOOTES_Visibility.Visibility")


Public Const pi = 3.14159265

'public void MoonCoordenatesAndBrightness_SelectDay(short Year,short Month,short Day,double Hour, double objRA, objDEC, ref double bright, ref double distance)

'Este metodo calcula el tanto por 1 de la luna que está iluminado y la distancia en Grados Sexagesimales que hay respecto a un objeto.

'Como entrada de esta función debemos indicar la fecha a la que queremos realizar la observacion. El dato de fecha hay que introducirlo 
'numéricamente. Para las coordenadas del objeto, objRA, son las coordenadas RA en Horas y objDEC son las coordenadas DEC en Grados Sexagesimales.
 

'Declaramos las variables necesarias para introducir al método

Public year1,month1,day1,hour1,objRA, objDEC, distance, bright


'Asignamos valores para probar el método

objRA= 2.0
objDEC = 6.0
year1 = 2017
month1 = 11
day1 = 30
hour1 = 20 

'Asignamos valores nulos a las variables que se calcularán

distance = 0
bright = 0

'Llamamos al método

myObj.MoonCoordenatesAndBrightness_SelectDay year1, month1, day1, hour1, objRA, objDEC, bright, distance

'Sacamos por pantalla los valores del brillo y la distancia.

Wscript.Echo "brillo de la luna:  ", bright
Wscript.Echo "Distancia respecto a la luna: ", distance