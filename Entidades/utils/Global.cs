﻿using System;
using System.IO;
using System.Xml;

namespace Entidades.utils
{
    public class Global
    {
        public static XmlDocument XmlDocument { get; set; }
        public static string ExcelFile { get; set; }

        #region RUTAS ESTÁTICAS
        public static string RutaAppExe = Environment.CurrentDirectory;
        public static string RutaDirectorioData = Path.Combine(Environment.CurrentDirectory, @"..\..\..\data\");
        #endregion

        #region RUTAS HTTP XML
        public static string RutaEnvioPruebas = "https://prewww1.aeat.es/wlpl/SSII-FACT/ws/fe/SiiFactFEV1SOAP";
        public const string SII = "https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroInformacion.xsd";
        public const string SII_R = "https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/RespuestaSuministro.xsd";
        public const string SII_LR = "https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroLR.xsd";
        public const string SOAPENV = "http://schemas.xmlsoap.org/soap/envelope/";
        #endregion

        #region RUTAS, DATO Y NOMBRE GUARDADO XML 
        public static string FechaGuardado { get; set; }
        public static string RutaGuardarXmlEnvio { get; set; }
        public static string RutaGuardarXmlRespuesta { get; set; }
        #endregion
    }
}
