
﻿using System.Collections.Generic;

namespace Datos.XML
{
    public interface IConstructorXML
    {
        ConstructorXML EstructuraXML();
        ConstructorXML EstructuraCabeceraXML();
        void EstructuraFacturaXML(IEnumerable<Dictionary<int, dynamic>> diccionarioValores);
        void GuardarXML();
    }
}
