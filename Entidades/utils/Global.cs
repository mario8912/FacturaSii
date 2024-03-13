﻿using System;
using System.IO;
using System.Xml;

namespace Entidades.utils
{
    public class Global
    {
        public static string RutaApplicacion = Path.Combine(Environment.CurrentDirectory, @"..\..\data");
        public static string RutaGuardarXml = Path.Combine(Environment.CurrentDirectory, @"..\..\..\data\nuevo.xml");

        public const string SII = "https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroInformacion.xsd";
        public const string SII_LR = "https://www2.agenciatributaria.gob.es/static_files/common/internet/dep/aplicaciones/es/aeat/ssii/fact/ws/SuministroLR.xsd";
        public const string SOAPENV = "http://schemas.xmlsoap.org/soap/envelope/";
        public static XmlDocument XmlDocument { get; set; }
        public static string ExcelFile { get; set; }
    }

    public class EventoProgreso
    {
        private int _progreso = 0;
        public int ValorMaximoBarraProgreso { get; set; }   

        public event EventHandler<int> ProgresoCambiado;

        public void AumentarProgreso()
        {
            _progreso++;
            OnProgresoCambiado();
        }

        protected virtual void OnProgresoCambiado()
        {
            ProgresoCambiado?.Invoke(this, _progreso);
        }
    }
}
