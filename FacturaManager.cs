using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FacturasXMLAExcelManager
{
    internal class FacturaManager
    {

        private String numero;
        private String rut;
        private String proveedor;
        private String moneda;
        private String tipoCambio;
        private String FechaDoc;
        private String total;
        private String estado;
        private String glosa;

        public FacturaManager()
        {
  
        }

        public FacturaManager(string numero, string rut, string proveedor, string moneda, string tipoCambio, string fechaDoc, string total, string estado, string glosa)
        {
            this.numero = numero;
            this.rut = rut;
            this.proveedor = proveedor;
            this.moneda = moneda;
            this.tipoCambio = tipoCambio;
            FechaDoc = fechaDoc;
            this.total = total;
            this.estado = estado;
            this.glosa = glosa;
        }

        public string Numero { get => numero; set => numero = value; }
        public string Rut { get => rut; set => rut = value; }
        public string Proveedor { get => proveedor; set => proveedor = value; }
        public string Moneda { get => moneda; set => moneda = value; }
        public string TipoCambio { get => tipoCambio; set => tipoCambio = value; }
        public string FechaDoc1 { get => FechaDoc; set => FechaDoc = value; }
        public string Total { get => total; set => total = value; }
        public string Estado { get => estado; set => estado = value; }
        public string Glosa { get => glosa; set => glosa = value; }
    }
}
