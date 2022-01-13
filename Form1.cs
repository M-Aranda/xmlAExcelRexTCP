using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace FacturasXMLAExcelManager
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //El excel a subir es el del formato de importación de documentos históricos con detalle
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<Factura> facturas = new List<Factura>();

            string sFileName = "";

            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "All Files (*.*)|*.*";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            string[] arrAllFiles= new string[] {};

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                sFileName = choofdlog.FileName;
                arrAllFiles = choofdlog.FileNames; //used when Multiselect = true           
            }

            Boolean variasFacturas = true;
   


            if (variasFacturas == false)
            {
    
                String URLString = sFileName; 

                Factura f = new Factura();


                f.TipoDeDocumento = getValue("TipoDTE", sFileName);


                //fijarse con el SII
                switch (f.TipoDeDocumento)
                {
                    case "33":
                        f.TipoDeDocumento = "FACE";
                        break;
                    case "34":
                        f.TipoDeDocumento = "FCEE";
                        break;
                    case "61":
                        f.TipoDeDocumento = "NCCE";
                        break;
                    default:
                        break;
                }


                //Las fechas son en formato dd/mm/yyyy
                f.NumeroDelDocumento = getValue("Folio", sFileName);
                f.FechaDeDocumento = getValue("FchEmis", sFileName);
                f.FechaContableDeDocumento = getValue("FchEmis", sFileName);//que es la fecha de cancelacion?
                f.FechaDeVencimientoDeDocumento = getValue("FchVenc", sFileName);
                f.CodigoDeUnidadDeNegocio = "0"; //getValue("Folio", sFileName);
                f.RutCliente = getValue("RUTRecep", sFileName);// o sea nosotoros getValue("TipoDTE", sFileName);
                f.DireccionDelCliente = getValue("DirRecep", sFileName);
                f.RutFacturador = getValue("RUTEmisor", sFileName);
                f.CodigoVendedor = "";// getValue("TipoDTE", sFileName);
                f.CodigoComisionista = "";// getValue("Folio", sFileName);
                f.Probabilidad = "";// getValue("Folio", sFileName);
                f.ListaPrecio = "";//getValue("TipoDTE", sFileName);
                f.PlazoPago = "";//getValue("Folio", sFileName);
                f.MonedaDelDocumento = "CLP";//getValue("Folio", sFileName);
                f.TasaDeCambio = "";// getValue("TipoDTE", sFileName);
                f.MontoAfecto = getValue("MntNeto", sFileName);

                f.MontoExento = getValue("MontoImp", sFileName);
                //ojo con el monto exento

                //MontoImp?
                //Si la factura es electronica no afecta, el MontoImp = MontoExento


                f.MontoIva = getValue("IVA", sFileName);
                f.MontoImpuestosEspecificos = "";//getValue("Folio", sFileName);
                f.MontoIvaRetenido = "";//getValue("Folio", sFileName);
                f.MontoImpuestosRetenidos = "";// getValue("TipoDTE", sFileName);
                f.TipoDeDescuentoGlobal = "";//getValue("Folio", sFileName);
                f.DescuentoGlobal = "";//getValue("Folio", sFileName);
                f.TotalDelDocumento = getValue("MntTotal", sFileName);
               
                f.DeudaPendiente = "0";//getValue("Folio", sFileName);
                f.TipoDocReferencia = "";//getValue("Folio", sFileName);
                f.NumDocReferencia = "";//getValue("Folio", sFileName);
                f.FechaDocReferencia = "";//getValue("Folio", sFileName);
                f.CodigoDelProducto = "410103";//getValue("TipoDTE", sFileName);
                f.Cantidad = "1"; getValue("Folio", sFileName);
                f.Unidad = "S.U.M"; //getValue("Folio", sFileName);
                f.PrecioUnitario = getValue("MntTotal", sFileName);
                f.MonedaDelDetalle = "CLP";//getValue("CLP", sFileName);
                f.TasaDeCambio2 = "1";//getValue("TipoDTE", sFileName);
                f.NumeroDeSerie = "";//getValue("Folio", sFileName);
                f.NumeroDeLote = "";//getValue("Folio", sFileName);
                f.FechaDeVencimiento = "";// getValue("Folio", sFileName);
                f.CentroDeCostos = getValue("CmnaDest", sFileName);
                f.TipoDeDescuento = "";//getValue("TipoDTE", sFileName);
                f.Descuento = "";//getValue("Folio", sFileName);
                f.Ubicacion = "";//getValue("Folio", sFileName);
                f.Bodega = "";//getValue("TipoDTE", sFileName);
                f.Concepto1 = "";//getValue("Folio", sFileName);
                f.Concepto2 = "";//getValue("Folio", sFileName);
                f.Concepto3 = "";//getValue("TipoDTE", sFileName);
                f.Concepto4 = "";//getValue("Folio", sFileName);
                f.Descripcion = "";//getValue("Folio", sFileName);
                f.DescripcionAdicional = "";//getValue("Folio", sFileName);
                f.Stock = "0";//getValue("Folio", sFileName);
                f.Comentario11 = "";// getValue("TipoDTE", sFileName);
                f.Comentario21 = "";//getValue("Folio", sFileName);
                f.Comentario31 = "";//getValue("Folio", sFileName);
                f.Comentario41 = "";//getValue("Folio", sFileName);
                f.Comentario51 = "";//getValue("Folio", sFileName);
                f.CodigoImpuestoEspecifico1 = "";// getValue("TipoDTE", sFileName);
                f.MontoImpuestoEspecifico1 = "";// getValue("Folio", sFileName);
                f.CodigoImpuestoEspecifico2 = "";//getValue("Folio", sFileName);
                f.MontoImpuestoEspecifico2 = "";//getValue("Folio", sFileName);
                f.Modalidad = "";//getValue("Folio", sFileName);
                f.Glosa = "FALTANTE";//getValue("Folio", sFileName);
                f.Referencia = "";//getValue("Folio", sFileName);
                f.FechaDeComprometida = "";//getValue("Folio", sFileName);
                f.PorcentajeCEEC = "";//getValue("Folio", sFileName);
                f.ImpuestoLey18211 = "";//getValue("Folio", sFileName);
                f.IvaLey18211 = "";//getValue("Folio", sFileName);
                f.CodigoKitFlexible = "";//getValue("Folio", sFileName);
                f.AjusteIva = "";//getValue("Folio", sFileName);







                facturas.Add(f);

            }

            else
            {
                foreach (var item in arrAllFiles)
                {
                     sFileName = item; 
                   // XmlTextReader reader = new XmlTextReader(URLString);

                    Factura f = new Factura();

                    f.MontoAfecto = "0";
                    f.MontoExento = "0";
                    f.MontoIva = "0";
                    f.TotalDelDocumento = "0";
                    


                    f.TipoDeDocumento = getValue("TipoDTE", sFileName);
                    

                    //fijarse con el SII
                    switch (f.TipoDeDocumento)
                    {
                        case "33":
                            f.TipoDeDocumento = "FACE";
                            break;
                        case "34":
                            f.TipoDeDocumento = "FCEE";
                            break;
                        case "61":
                            f.TipoDeDocumento = "NCCE";
                            break;
                        default:
                            break;
                    }

                    
                    //Las fechas son en formato dd/mm/yyyy
                    f.NumeroDelDocumento = getValue("Folio", sFileName);
                    f.FechaDeDocumento = convertirAFechaValida(getValue("FchEmis", sFileName));
                    f.FechaContableDeDocumento = convertirAFechaValida(getValue("FchEmis", sFileName));//que es la fecha de cancelacion?
                    f.FechaDeVencimientoDeDocumento = convertirAFechaValida(getValue("FchEmis", sFileName));//convertirAFechaValida(getValue("FchVenc", sFileName));// fecha de vencimiento debe ser igual o mayor a fecha de emision
                    f.CodigoDeUnidadDeNegocio = "1"; //getValue("Folio", sFileName);
                    f.RutCliente = getValue("RUTEmisor", sFileName);
                    f.DireccionDelCliente = "Casa Matriz"; //getValue("DirRecep", sFileName);
                    f.RutFacturador = "";//getValue("RUTEmisor", sFileName);//getValue("RUTRecep", sFileName);
                    f.CodigoVendedor = "";// getValue("TipoDTE", sFileName);
                    f.CodigoComisionista = "";// getValue("Folio", sFileName);
                    f.Probabilidad = "";// getValue("Folio", sFileName);
                    f.ListaPrecio = "";//getValue("TipoDTE", sFileName);
                    f.PlazoPago = "P01";//getValue("Folio", sFileName); codigo de plazo pago?
                    f.MonedaDelDocumento = "CLP";//getValue("Folio", sFileName);
                    f.TasaDeCambio = "";// getValue("TipoDTE", sFileName);
                    f.MontoAfecto = getValue("MntNeto", sFileName);


                    f.MontoExento = getValue("MontoImp", sFileName);

                    if (f.TipoDeDocumento=="FCEE")
                    {f.MontoExento= getValue("MntExe", sFileName);

                    }
                    //ojo con el monto exento

                    //MontoImp?
                    //Si la factura es electronica  afecta, el MontoImp = MontoExento
                    //Si la factura es electronica no afecta o exenta MntExe = MontoExento


                    f.MontoIva = getValue("IVA", sFileName);
                    f.MontoImpuestosEspecificos = "";//getValue("Folio", sFileName);
                    f.MontoIvaRetenido = "";//getValue("Folio", sFileName);
                    f.MontoImpuestosRetenidos = "";// getValue("TipoDTE", sFileName);
                    f.TipoDeDescuentoGlobal = "";//getValue("Folio", sFileName);
                    f.DescuentoGlobal = "";//getValue("Folio", sFileName);
                    f.TotalDelDocumento = getValue("MntTotal", sFileName);
                    f.DeudaPendiente = "0";//getValue("MntTotal", sFileName);
                    f.TipoDocReferencia = "";//getValue("Folio", sFileName);
                    f.NumDocReferencia = "";//getValue("Folio", sFileName);
                    f.FechaDocReferencia = "";//getValue("Folio", sFileName);
                    f.CodigoDelProducto = "410103";//getValue("TipoDTE", sFileName);
                    f.Cantidad = "1"; getValue("Folio", sFileName);
                    f.Unidad = "S.U.M"; //getValue("Folio", sFileName);
                    f.PrecioUnitario = getValue("MntTotal", sFileName);
                    f.MonedaDelDetalle = "CLP";
                    f.TasaDeCambio2 = "1";//getValue("TipoDTE", sFileName);
                    f.NumeroDeSerie = "";//getValue("Folio", sFileName);
                    f.NumeroDeLote = "";//getValue("Folio", sFileName);
                    f.FechaDeVencimiento = "";// getValue("Folio", sFileName);
                    f.CentroDeCostos = "204"; //getValue("CmnaDest", sFileName); //Este es el centro de costos
                    f.TipoDeDescuento = "";//getValue("TipoDTE", sFileName);
                    f.Descuento = "";//getValue("Folio", sFileName);
                    f.Ubicacion = "";//getValue("Folio", sFileName);
                    f.Bodega = "";//getValue("TipoDTE", sFileName);
                    f.Concepto1 = "";//getValue("Folio", sFileName);
                    f.Concepto2 = "";//getValue("Folio", sFileName);
                    f.Concepto3 = "";//getValue("TipoDTE", sFileName);
                    f.Concepto4 = "";//getValue("Folio", sFileName);
                    f.Descripcion = "";//getValue("Folio", sFileName);
                    f.DescripcionAdicional = "";//getValue("Folio", sFileName);
                    f.Stock = "0";//getValue("Folio", sFileName);
                    f.Comentario11 = "";// getValue("TipoDTE", sFileName);
                    f.Comentario21 = "";//getValue("Folio", sFileName);
                    f.Comentario31 = "";//getValue("Folio", sFileName);
                    f.Comentario41 = "";//getValue("Folio", sFileName);
                    f.Comentario51 = "";//getValue("Folio", sFileName);
                    f.CodigoImpuestoEspecifico1 = "";// getValue("TipoDTE", sFileName);
                    f.MontoImpuestoEspecifico1 = "";// getValue("Folio", sFileName);
                    f.CodigoImpuestoEspecifico2 = "";//getValue("Folio", sFileName);
                    f.MontoImpuestoEspecifico2 = "";//getValue("Folio", sFileName);
                    f.Modalidad = "";//getValue("Folio", sFileName);
                    f.Glosa = "FALTANTE";//getValue("Folio", sFileName);
                    f.Referencia = "";//getValue("Folio", sFileName);
                    f.FechaDeComprometida = "";//getValue("Folio", sFileName);
                    f.PorcentajeCEEC = "";//getValue("Folio", sFileName);
                    f.ImpuestoLey18211 = "";//getValue("Folio", sFileName);
                    f.IvaLey18211 = "";//getValue("Folio", sFileName);
                    f.CodigoKitFlexible = "";//getValue("Folio", sFileName);
                    f.AjusteIva = "";//getValue("Folio", sFileName);




                    facturas.Add(f);

                }


            }



            var archivo = new FileInfo(@"C:\Users\Chelo\Desktop\excelsDeFacturas\FacturasEnExcel.xlsx");

            SaveExcelFile(facturas, archivo);

            MessageBox.Show("Se creó el archivo");



        }


    

        private static async Task SaveExcelFile(List<Factura> facturas, FileInfo file)
        {
            var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("Facturas");

            var range = ws.Cells["A1"].LoadFromCollection(facturas, true);

            range.AutoFitColumns();

            await package.SaveAsync();
        }

        public string getValue(string _clave, String pathDelArchivo)
        {
            XmlTextReader textReader = new XmlTextReader(pathDelArchivo);
            textReader.Read();
            while (textReader.Read())
            {
                textReader.MoveToElement();
                if (textReader.Name == _clave)
                {
                    textReader.Read();
                    if (textReader.Value.ToString().Trim() != "")
                    {
                        string z = textReader.Value.ToString().Trim();
                        textReader.Close();
                        return z;
                    }
                }
            }
            textReader.Close();
            return "";
        }

        public String convertirAFechaValida(String fechaAConvertir)
        {
            String fechaValida = "";


            string[] datos = fechaAConvertir.Split('-');

            fechaValida = datos[2]+"/"+datos[1]+"/" + datos[0];

            return fechaValida;


        }













    }




}
