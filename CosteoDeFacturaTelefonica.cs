using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FacturasXMLAExcelManager
{
    internal class CosteoDeFacturaTelefonica
    {

        private String numeroDeCelular;
        private String descripcion;
        private String numeroDeSerie;
        private String pin;
        private String puk;
        private String asignadoA;
        private String centro;
        private String activo;
        private String comentario;

        public CosteoDeFacturaTelefonica()
        {
        }

        public CosteoDeFacturaTelefonica(string numeroDeCelular, string descripcion, string numeroDeSerie, string pin, string puk, string asignadoA, string centro, string activo, string comentario)
        {
            this.NumeroDeCelular = numeroDeCelular;
            this.Descripcion = descripcion;
            this.NumeroDeSerie = numeroDeSerie;
            this.Pin = pin;
            this.Puk = puk;
            this.AsignadoA = asignadoA;
            this.Centro = centro;
            this.Activo = activo;
            this.Comentario = comentario;
        }

        public string NumeroDeCelular { get => numeroDeCelular; set => numeroDeCelular = value; }
        public string Descripcion { get => descripcion; set => descripcion = value; }
        public string NumeroDeSerie { get => numeroDeSerie; set => numeroDeSerie = value; }
        public string Pin { get => pin; set => pin = value; }
        public string Puk { get => puk; set => puk = value; }
        public string AsignadoA { get => asignadoA; set => asignadoA = value; }
        public string Centro { get => centro; set => centro = value; }
        public string Activo { get => activo; set => activo = value; }
        public string Comentario { get => comentario; set => comentario = value; }
    }
}
