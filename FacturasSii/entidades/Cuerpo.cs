using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FacturasSii.entidades
{
    internal class Cuerpo
    {
        public class RegistroLRFacturasEmitidas
        {
            public PeriodoLiquidacion PeriodoLiquidacion { get; set; }
            public IDFactura IDFactura { get; set; }
            public FacturaExpedida FacturaExpedida { get; set; }
        }
        public class PeriodoLiquidacion
        {
            public string Ejercicio { get; set; }
            public string Periodo { get; set; }
        }

        public class IDEmisorFactura
        {
            public string NIF { get; set; }
        }

        public class IDFactura
        {
            public IDEmisorFactura IDEmisorFactura { get; set; }
            public string NumSerieFacturaEmisor { get; set; }
            public string FechaExpedicionFacturaEmisor { get; set; }
        }

        public class Contraparte
        {
            public string NombreRazon { get; set; }
            public string NIF { get; set; }
        }

        public class DetalleIVA
        {
            public string TipoImpositivo { get; set; }
            public string BaseImponible { get; set; }
            public string CuotaRepercutida { get; set; }
        }

        public class DesgloseIVA
        {
            public DetalleIVA[] DetalleIVA { get; set; }
        }

        public class NoExenta
        {
            public string TipoNoExenta { get; set; }
            public DesgloseIVA DesgloseIVA { get; set; }
        }

        public class Sujeta
        {
            public NoExenta NoExenta { get; set; }
        }

        public class DesgloseFactura
        {
            public Sujeta Sujeta { get; set; }
        }

        public class TipoDesglose
        {
            public DesgloseFactura DesgloseFactura { get; set; }
        }

        public class FacturaExpedida
        {
            public string TipoFactura { get; set; }
            public string ClaveRegimenEspecialOTrascendencia { get; set; }
            public string ImporteTotal { get; set; }
            public string DescripcionOperacion { get; set; }
            public Contraparte Contraparte { get; set; }
            public TipoDesglose TipoDesglose { get; set; }
        }        
    }
}
