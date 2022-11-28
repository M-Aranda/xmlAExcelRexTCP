using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FacturasXMLAExcelManager
{
    internal class SubCosteoTelefonico
    {

        private String centro;
        private String valor;

        public SubCosteoTelefonico()
        {
        }

        public SubCosteoTelefonico(string centro, string valor)
        {
            this.Centro = centro;
            this.Valor = valor;
        }

        public string Centro { get => centro; set => centro = value; }
        public string Valor { get => valor; set => valor = value; }
    }
}
