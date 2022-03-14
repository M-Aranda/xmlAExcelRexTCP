using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FacturasXMLAExcelManager
{
    internal class Reporte
    {
        
        private String codigoDeCentro;
        private String nombreDeCentro;
        private String totalCosteadoAlCentro;


        public Reporte()
        {

        }

        public Reporte(string codigoDeCentro, string nombreDeCentro, string totalCosteadoAlCentro)
        {
            this.CodigoDeCentro = codigoDeCentro;
            this.NombreDeCentro = nombreDeCentro;
            this.TotalCosteadoAlCentro = totalCosteadoAlCentro;
        }

        public string CodigoDeCentro { get => codigoDeCentro; set => codigoDeCentro = value; }
        public string NombreDeCentro { get => nombreDeCentro; set => nombreDeCentro = value; }
        public string TotalCosteadoAlCentro { get => totalCosteadoAlCentro; set => totalCosteadoAlCentro = value; }
    }
}
