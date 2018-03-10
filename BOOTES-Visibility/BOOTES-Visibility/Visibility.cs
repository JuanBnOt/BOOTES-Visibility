using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace BOOTES_Visibility
{
    [ComVisible(true)]
    public class Visibility
        {

            [ComVisible(true)]
            public void MoonDistanceAndBrightness_SelectDay(short Year, short Month, short Day, double Hour, double objRA, double objDEC, ref double bright, ref double distance)
            {
                /*
                 * El objRA y objDEC deben estar en las unidades comunes para este tipo de coordenadas
                 * es decir RA en horas y DEC en grados, sin minutos y segundos ambos, deben ser por ejemplo
                 * objRA = 18.25 (horas) y objDEC = 76.758 (grados)
                 * 
                 * Devuelve la distancia a la luna y la luminosidad calculada en la fecha pasada por el parametro de entrada
                 * dt.
                 */

                //Inicialización de variables necesarias para calcular la posición de la luna

                ASCOM.Astrometry.Object3 moonObj3 = new ASCOM.Astrometry.Object3();
                ASCOM.Astrometry.Accuracy Accu;

                double JD = 0;
                double rc = 0;
                double moonDisUA = 0;
                double moonRA = 0;
                double moonDEC = 0;

                //Creamos los objetos ASCOM que contienen los metodos necesarios para el calculo de la
                //posicion de la luna, de la luminosidad, del Julian Date y conversiones de datos.

                ASCOM.Astrometry.NOVAS.NOVAS31 ApASCOM2 = new ASCOM.Astrometry.NOVAS.NOVAS31();
                ASCOM.Astrometry.AstroUtils.AstroUtils ApASCOM1 = new ASCOM.Astrometry.AstroUtils.AstroUtils();

                //Creamos el objeto ASCOM de la luna

                moonObj3.Name = "Moon";
                moonObj3.Type = ASCOM.Astrometry.ObjectType.MajorPlanetSunOrMoon;
                moonObj3.Number = ASCOM.Astrometry.Body.Moon;

                //Definimos la precision del calculo

                Accu = ASCOM.Astrometry.Accuracy.Full;

                //Calculamos el Julian Date para el dia pasado por entrada

                JD = ApASCOM2.JulianDate(Year, Month, Day, Hour);


                //Calculamos las coordenadas de la luna en RA (Horas) y DEC (grados)
                //rc es un parámetro de control que dice si la función se ha ejecutado correctamente

                rc = ApASCOM2.AppPlanet(JD, moonObj3, Accu, ref moonRA, ref moonDEC, ref moonDisUA);

                bright = ApASCOM1.MoonIllumination(JD);

                /* Pasamos al calculo de la distancia con el objeto que deseamos observar con el telescopio.
                 * Para ello tendremos que cambiar los datos de coordenadas de ambos objetos ya que RA esta
                 * en GRADOS y DEC en GRADOS SEXAGESIMALES. Con lo cual se haran las conversiones pertinentes para obtener radianes
                 * tras los cálculos en radianes se devolverá el resultado en GRADOS.
                 ****************************************************************************
                 */
                double moonRArad = moonRA * 15 * Math.PI / 180; //Conversión a  radianes (pi/180) y a grados (15)
                double moonDECrad = moonDEC * Math.PI / 180;
                double objRA1 = objRA * Math.PI / 180; //Conversión a radianes (pi/180)
                double objDEC1 = objDEC * Math.PI / 180; //Solo necesitamos conversión a radianes.

                //La formula para calcular la distancia ha sido obtenida de la página web http://aa.quae.nl/en/reken/afstanden.html (equation 11 from polar coordinates)
                //Devuelve el resultado de distance en Grados Sexagesimales

                //q es una variable auxiliar para poder hacer los calculos de manera mas ordenada.
                //q is a auxiliar variable just to calculate the distance in a clearer way.
              

                double q = Math.Pow(Math.Sin((0.5 * (objDEC1 - moonDECrad))), (double)2) + Math.Cos(moonDECrad) * Math.Cos(objDEC1) * Math.Pow(Math.Sin(0.5 * (objRA1 - moonRArad)), (double)2);
                double q2 = Math.Pow(Math.Sin((0.5 * (objDEC1 + moonDECrad))), (double)2) + Math.Cos(moonDECrad) * Math.Cos(objDEC1) * Math.Pow(Math.Cos(0.5 * (objRA1 - moonRArad)), (double)2);
                distance = Math.Abs(2 * Math.Atan(Math.Sqrt(q / q2)) * 180 / Math.PI);//el resultado de ATAN es radianes de ahí que se multiplique por la conversión para pasar a grados.
            }

            [ComVisible(true)]
            public void MoonDistanceAndData_SelectDay(short Year, short Month, short Day, double Hour, double objRA, double objDEC, ref double distance,  ref double bright, ref double moonRA, ref double moonDEC)
            {

                /*
                 * El objRA y objDEC deben estar en las unidades comunes para este tipo de coordenadas
                 * es decir RA en horas y DEC en grados, sin minutos y segundos ambos, deben ser por ejemplo
                 * objRA = 18.25 (horas) y objDEC = 76.758 (grados)
                 * 
                 * Devuelve la distancia a la luna y la luminosidad calculada en la fecha pasada por el parametro de entrada
                 * dt.
                 */

                //Inicialización de variables necesarias para calcular la posición de la luna

                ASCOM.Astrometry.Object3 moonObj3 = new ASCOM.Astrometry.Object3();
                ASCOM.Astrometry.Accuracy Accu;

                double JD = 0;
                double rc = 0;
                double moonDisUA = 0;
                

                //Creamos los objetos ASCOM que contienen los metodos necesarios para el calculo de la
                //posicion de la luna, de la luminosidad, del Julian Date y conversiones de datos.

                ASCOM.Astrometry.NOVAS.NOVAS31 ApASCOM2 = new ASCOM.Astrometry.NOVAS.NOVAS31();
                ASCOM.Astrometry.AstroUtils.AstroUtils ApASCOM1 = new ASCOM.Astrometry.AstroUtils.AstroUtils();

                //Creamos el objeto ASCOM de la luna

                moonObj3.Name = "Moon";
                moonObj3.Type = ASCOM.Astrometry.ObjectType.MajorPlanetSunOrMoon;
                moonObj3.Number = ASCOM.Astrometry.Body.Moon;

                //Definimos la precision del calculo

                Accu = ASCOM.Astrometry.Accuracy.Full;

                //Calculamos el Julian Date para el dia pasado por entrada

                JD = ApASCOM2.JulianDate(Year, Month, Day, Hour);


                //Calculamos las coordenadas de la luna en RA (Horas) y DEC (grados)
                //rc es un parámetro de control que dice si la función se ha ejecutado correctamente

                rc = ApASCOM2.AppPlanet(JD, moonObj3, Accu, ref moonRA, ref moonDEC, ref moonDisUA);

                //Calculamos el brillo de la luna en tanto por 1

                bright = ApASCOM1.MoonIllumination(JD);

                /* Pasamos al calculo de la distancia con el objeto que deseamos observar con el telescopio.
                 * Para ello tendremos que cambiar los datos de coordenadas de ambos objetos ya que RA esta
                 * en GRADOS y DEC en GRADOS SEXAGESIMALES. Con lo cual se haran las conversiones pertinentes para obtener radianes
                 * tras los cálculos en radianes se devolverá el resultado en GRADOS.
                 ****************************************************************************
                 */
                moonRA = moonRA * 15 * Math.PI / 180; //Conversión a  radianes (pi/180) y a grados (15)
                moonDEC = moonDEC * Math.PI / 180;
                double objRA1 = objRA * Math.PI / 180; //Conversión a radianes (pi/180)
                double objDEC1 = objDEC * Math.PI / 180; //Solo necesitamos conversión a radianes.

                //La formula para calcular la distancia ha sido obtenida de la página web http://aa.quae.nl/en/reken/afstanden.html (equation 11 from polar coordinates)
                //Devuelve el resultado de distance en Grados Sexagesimales

                //q es una variable auxiliar para poder hacer los calculos de manera mas ordenada.
                //q is a auxiliar variable just to calculate the distance in a clearer way.


                double q = Math.Pow(Math.Sin((0.5 * (objDEC1 - moonDEC))), (double)2) + Math.Cos(moonDEC) * Math.Cos(objDEC1) * Math.Pow(Math.Sin(0.5 * (objRA1 - moonRA)), (double)2);
                double q2 = Math.Pow(Math.Sin((0.5 * (objDEC1 + moonDEC))), (double)2) + Math.Cos(moonDEC) * Math.Cos(objDEC1) * Math.Pow(Math.Cos(0.5 * (objRA1 - moonRA)), (double)2);
                distance = Math.Abs(2 * Math.Atan(Math.Sqrt(q / q2)) * 180 / Math.PI);//el resultado de ATAN es radianes de ahí que se multiplique por la conversión para pasar a grados.

                //Devolvemos a grados el moonRA y moonDEC

                moonRA =  Math.Round(moonRA * 180 / Math.PI,12); //Conversión a  grados (180/pi)
                moonDEC = Math.Round(moonDEC * 180 / Math.PI,12);
            }

        }


    
}