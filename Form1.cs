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
using System.Xml.Linq;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;
using System.Data;
using System.Data.SqlClient;
using Windows.Storage;

namespace FacturasXMLAExcelManager
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //El excel a subir es el del formato de importación de documentos contables con detalle
            //estos botones son para las otras funciones
            button2.Visible = false;
            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;

        }

        private void button1_Click(object sender, EventArgs e)
        {

            //excelAPartirDeXML
            List<Factura> facturas = new List<Factura>();
            List<FacturaNCCE> facturasNCCE = new List<FacturaNCCE>();
            

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
                f.TipoDeDocumento = determinarTipoDeDocumento(f.TipoDeDocumento);


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


                //Hay que ingresar en mantenedores, importador de datos, datos a importar,
                //Documentos con detalles (Contabilizado)
                //hay 3 escenarios: si es factura exenta de iva, todo es 0, menos el campo de exento


                //si es factura electronica


                //MontoImp?
                //Si la factura es electronica no afecta, el MontoImp = MontoExento


                f.MontoIva = getValue("IVA", sFileName);
                f.MontoImpuestosEspecificos = "";//getValue("Folio", sFileName);
                f.MontoIvaRetenido = "";//getValue("Folio", sFileName);
                f.MontoImpuestosRetenidos = "";// getValue("TipoDTE", sFileName);
                f.TipoDeDescuentoGlobal = "";//getValue("Folio", sFileName);
                f.DescuentoGlobal = "";//getValue("Folio", sFileName);
                f.TotalDelDocumento = getValue("MntTotal", sFileName);
               
                f.DeudaPendiente = getValue("MntTotal", sFileName);// deuda pendiente siempre es igual al monto total
                f.TipoDocReferencia = "";//getValue("Folio", sFileName);
                f.NumDocReferencia = "";//getValue("Folio", sFileName);
                f.FechaDocReferencia = "";//getValue("Folio", sFileName);
                f.CodigoDelProducto = "410103";//getValue("TipoDTE", sFileName);
                //El codigo del producto va a depender de lo que sea el item
                f.Cantidad = "1"; getValue("Folio", sFileName);
                f.Unidad = "S.U.M"; //getValue("Folio", sFileName);
                f.PrecioUnitario = getValue("MntTotal", sFileName);
                f.MonedaDelDetalle = "CLP";//getValue("CLP", sFileName);
                f.TasaDeCambio2 = "1";//getValue("TipoDTE", sFileName);
                f.NumeroDeSerie = "";//getValue("Folio", sFileName);
                f.NumeroDeLote = "";//getValue("Folio", sFileName);
                f.FechaDeVencimiento = "";// getValue("Folio", sFileName);
                f.CentroDeCostos = getValue("DirOrigen", sFileName);//no es CmnaDest el dato que me da la dirección, es DirOrigen
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


                    // los ruts de CCU son:
                    // 91041000-8
                    //96989120-4
                    //99501760-1
                    //99554560-8
                    //99586280-8
                  


                     sFileName = item; 
                   // XmlTextReader reader = new XmlTextReader(URLString);

                    Factura f = new Factura();

                    f.MontoAfecto = "0";
                    f.MontoExento = "0";
                    f.MontoIva = "0";
                    f.TotalDelDocumento = "0";

                    //Hay que ver como leer los detalles

                    XmlTextReader textReader = new XmlTextReader(sFileName);
                    textReader.Read();
                    
                    //List<DetalleDeFactura> detalles = new List<DetalleDeFactura>();
                    //DetalleDeFactura detalle = new DetalleDeFactura();

                    List<String> datos = new List<String>();


                    List<List<String>> datosDeDatos = new List<List<string>>();

                    while (textReader.Read())
                    {
                       // Console.WriteLine("esto es un nodo");
                        String nombreItem = "";
                        //String cantidadItem = "";
                        //String unmdItem = "";
                        //String prcItem = "";
                        //String codImpAdic = "";
                        //String montoItem = "";

                        //textReader.MoveToElement();
                        //if (textReader.Name == "NmbItem")
                        //{
                        //    textReader.Read();
                        //    if (textReader.Value.ToString().Trim() != "")
                        //    {
                        //        detalle.NmbItem = textReader.Value.ToString();

                        //        nombreItem = detalle.NmbItem;
                        //        Console.WriteLine(nombreItem);

                                
                        //    }
                          

                        //}




                        //detalle = new DetalleDeFactura();
                        //detalle.NmbItem= nombreItem;
                        ////detalle.QtyItem= cantidadItem;
                        ////detalle.UnmdItem= unmdItem ;
                        ////detalle.PrcItem= prcItem ;
                        ////detalle.CodImpAdic= codImpAdic ;
                        ////detalle.MontoItem= montoItem ;

                        //detalles.Add(detalle);

                        

                      

                    }

          

                   

                    //foreach (var i in detalles)
                    //{
                    //    //Console.WriteLine(i.NmbItem);

                    //}



                    f.TipoDeDocumento = getValue("TipoDTE", sFileName);


                    f.TipoDeDocumento = determinarTipoDeDocumento(f.TipoDeDocumento);


                    //Las fechas son en formato dd/mm/yyyy
                    f.NumeroDelDocumento = getValue("Folio", sFileName);
                    f.FechaDeDocumento = convertirAFechaValida(getValue("FchEmis", sFileName));
                    f.FechaContableDeDocumento = convertirAFechaValida(getValue("FchEmis", sFileName));//que es la fecha de cancelacion?
                    f.FechaDeVencimientoDeDocumento = convertirAFechaValida(getValue("FchEmis", sFileName));//convertirAFechaValida(getValue("FchVenc", sFileName));// fecha de vencimiento debe ser igual o mayor a fecha de emision

                    DateTime now = DateTime.Now;
                    f.FechaContableDeDocumento = convertirAFechaValidaDesdeTranstecnia(Convert.ToString(now.Date));//"dia actual"

                    f.FechaDeDocumento = f.FechaContableDeDocumento;
                    f.FechaDeVencimientoDeDocumento = f.FechaContableDeDocumento;


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
                    f.DeudaPendiente = getValue("MntTotal", sFileName);
                    f.TipoDocReferencia = "";//getValue("Folio", sFileName);
                    f.NumDocReferencia = "";//getValue("Folio", sFileName);
                    f.FechaDocReferencia = "";//getValue("Folio", sFileName);

                    //El codigo del producto varia, puede ser uno de los siguientes:

                    /*
                     
                    420710
                    420724E

 

        110904	Impuesto Específico	Generico		
		310101	Fletes Interplantas	Servicio	
		310201	Fletes Emprendedores	Servicio	
		310202	Recuperacion de Gastos EMP	Servicio	
		310203	Arriendo Vehiculos	Servicio	
		310301	Fletes Tradicional	Servicio	
		310302	Recuperacion de Gastos	Servicio	
		410101	Contratista PSCP	Gasto	
		410102	Fletes Acarreo	Gasto	
		410103	Fletes de terceros	Gasto	
					
	    410104	Petroleo	Gasto		
		410105	Servicios de Terceros	Gasto	
		410105E	Transporte de Pasajeros	Gasto	
		410106	Arriendo Vehiculos	Gasto	
		410107	Arriendo Leasing	Gasto	
		410108	Fletes emprendedores	Gasto	
		410109	Cuota Arriendo Vehiculo (Socio)	Gasto	
		410110	Arriendo Vehiculo (Socio)	Gasto	
		420102	Honorarios	Gasto	
		420501	Contrato Mantenimiento	Gasto	
	    420502	Mantención Equipamiento (Bs.muebles	Gasto		
		420503	Mant. Inmuebles	Gasto	
		420504	Lubricantes	Gasto	
		420505	Insumos y Repuestos	Gasto	
		420506	Mantención Neumáticos	Gasto	
		420507	Otras Mantenciones	Gasto	
		420701	Arriendo de Oficinas	Gasto	
		420702	Gastos Comunes	Gasto	
		420703	Insumos de Aseo	Gasto	
		420704	Insumos de Oficinas	Gasto	
	    420705	Soporte Computacional	Gasto		
		420706	Energia Electrica	Gasto	
		420707	Agua	Gasto	
		420708	Gas	Gasto	
		420709	Telefonia y Comunicación	Gasto	
		420710	Gastos de Supermercados	Gasto	
		420711	Servicios de Correos	Gasto	
		420712	Servicios de Vigilancia	Gasto	
		420713	Gastos Notariales	Gasto	
		420714	Sanitizaciones	Gasto	
	    420715	Evaluaciones Medicas	Gasto		
		420716	Ropa de Trabajo del Personal	Gasto	
		420717	Arriendo de Maquinarias y Equipos	Gasto	
		420718	Seguros Camionetas, Camiones y otros	Gasto	
		420719	Seguros Oficinas	Gasto	
		420720	Patentes Municipales	Gasto	
		420721	Patentes y Permisos de Vehiculos	Gasto	
		420722	Revisiones Tecnicas	Gasto	
		420723	Infracciones de Transito	Gasto	
		420724	Faltantes	Gasto	
	    420724E	Faltantes exento	Gasto		
		420724ESP	Faltantes con impuesto específico	Gasto	
		420725	Gastos Varios	Gasto	
		420726	Honorarios de Auditoria Externa	Gasto	
		420727	Asesoria Legal	Gasto	
		420804	Gastos Bancarios	Gasto	

*/





                    f.CodigoDelProducto = "420710";//getValue("TipoDTE", sFileName);
                    f.Cantidad = "1"; //getValue("Folio", sFileName);
                    f.Unidad = "S.U.M"; //getValue("Folio", sFileName);
                    f.PrecioUnitario = getValue("MntNeto", sFileName);
                    f.MonedaDelDetalle = "CLP";
                    f.TasaDeCambio2 = "1";//getValue("TipoDTE", sFileName);
                    f.NumeroDeSerie = "";//getValue("Folio", sFileName);
                    f.NumeroDeLote = "";//getValue("Folio", sFileName);
                    f.FechaDeVencimiento = "";// getValue("Folio", sFileName);
                    f.CentroDeCostos = getValue("DirOrigen", sFileName);//no es CmnaDest el dato que me da la dirección, es DirOrigen

                    //determinar a donde se costea
                    //los codigos de centros de costo son (numero de la izquierda: TCP, numero de la derecha: PSCP): 

                    //203 / 303   Administracion
                    //204 / 304   Interplantas
                    //208 / 308   Emprendedores
                    //205 / 305   Illapel
                    //207 / 307   San Antonio
                    //200 / 300   Melipilla
                    //206 / 306   Santiago
                    //201 / 301   Rancagua
                    //202 / 302   Curico


                    //si el rut del receptor es 78462150-2, el costeo es para TCP, si
                    //es 78877610-1, es para PSCP
                    String rutDeReceptor= getValue("RutReceptor", sFileName);
                    Boolean esPSCP = false;
                    if (rutDeReceptor== "78877610-1")
                    {
                        esPSCP = true;
                    }

                    f.CentroDeCostos = determinarCentroDeCosto(f.CentroDeCostos, esPSCP);

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
                    if (f.TipoDeDocumento == "NCCE")
                    {
                        f.Modalidad = "3";
                    }

                    f.Glosa = "Factura de compra";//getValue("Folio", sFileName);
                    f.Referencia = "";//getValue("Folio", sFileName);
                    f.FechaDeComprometida = "";//getValue("Folio", sFileName);
                    f.PorcentajeCEEC = "";//getValue("Folio", sFileName);
                    f.ImpuestoLey18211 = "";//getValue("Folio", sFileName);
                    f.IvaLey18211 = "";//getValue("Folio", sFileName);
                    f.CodigoKitFlexible = "";//getValue("Folio", sFileName);
                    f.AjusteIva = "";//getValue("Folio", sFileName);


                    f.CodigoDelProducto = "420710";
                    f.PrecioUnitario = getValue("MntNeto", sFileName);
                    if (String.IsNullOrEmpty(f.MontoExento) == true)
                    {
                        f.MontoExento = "0";
                    }

                    String exentoAntes = f.MontoExento;
                    Boolean esFacturaDeEnvases = false;

                    //si es una factura de envase, el exento y el total debiesen ser el mismo
                    //if (f.NumeroDelDocumento == "128012803")
                    //{
                    //    MessageBox.Show(f.MontoAfecto);
                    //    MessageBox.Show(f.MontoExento);
                    //    MessageBox.Show(f.MontoIva);
                    //}
                    if (f.MontoAfecto == "0" && f.MontoExento == "0" && f.MontoIva == "0" && f.TotalDelDocumento != "0")
                    {
                        
                        f.CodigoDelProducto = "420724E";
                        f.PrecioUnitario = f.TotalDelDocumento;
                        f.MontoExento = f.TotalDelDocumento;
                        f.TipoDeDocumento = "FCEE";
                        esFacturaDeEnvases = true;
                    }


                    if (f.RutCliente != "91041000-8" & f.RutCliente != "96989120-4" & f.RutCliente != "99501760-1" & f.RutCliente != "99554560-8" & f.RutCliente != "99586280-8")
                    {
                        f.CodigoDelProducto = "110804";
                        f.CentroDeCostos = "209";
                    }

                    if (f.TipoDeDocumento != "NCCE")
                    {

                        if (f.MontoIva == "0" && f.TipoDeDocumento=="FACE")
                        {
                            f.TipoDeDocumento = "FCEE";
                            f.CodigoDelProducto = "420724E";
                            
                        }


                        facturas.Add(f);
                    }
                    else
                    {

                        FacturaNCCE facNCCE = convertirFacturaANCCE(f);
                        facNCCE.TipoDeDocumentoDeOrigen = determinarTipoDeDocumento(getValue("TpoDocRef",sFileName));
                        facNCCE.NumeroDocumentoDeOrigen = getValue("FolioRef", sFileName);

           

                        facNCCE.FechaDeContableDeDocumento = convertirAFechaValida2(facNCCE.FechaDeContableDeDocumento);
                        facNCCE.FechaDeDocumento = facNCCE.FechaDeContableDeDocumento;
                        facNCCE.FechaDeVencimientoDeDocumento = facNCCE.FechaDeContableDeDocumento;

                        if (facNCCE.RutCliente != "91041000-8" & facNCCE.RutCliente != "96989120-4" & facNCCE.RutCliente != "99501760-1" & facNCCE.RutCliente != "99554560-8" & facNCCE.RutCliente != "99586280-8")
                        {
                            facNCCE.CodigoDelProducto = "110804";
                            facNCCE.CentroDeCostos = "209";
                        }

                        //420724E
                        facturasNCCE.Add(facNCCE);

                        // si el exento es distinto a 0, las notas de credito tienen que tener el mismo traqtamiento con las face cuyo exento es superior a 0
                        if(facNCCE.MontoExento!="0")
                        {
                            FacturaNCCE facNCCE2 = new FacturaNCCE();
                            facNCCE2 = convertirFacturaANCCE(f);
                            facNCCE2.PrecioUnitario = facNCCE2.MontoExento;
                            facNCCE2.CodigoDelProducto = "420724E";

                            facNCCE2.TipoDeDocumentoDeOrigen = facNCCE.TipoDeDocumentoDeOrigen;
                            facNCCE2.NumeroDocumentoDeOrigen = facNCCE.NumeroDocumentoDeOrigen;

                            facNCCE2.FechaDeContableDeDocumento = convertirAFechaValida2(facNCCE2.FechaDeContableDeDocumento);
                            facNCCE2.FechaDeDocumento = facNCCE2.FechaDeContableDeDocumento;
                            facNCCE2.FechaDeVencimientoDeDocumento = facNCCE2.FechaDeContableDeDocumento;

                            if (facNCCE2.RutCliente != "91041000-8" & facNCCE2.RutCliente != "96989120-4" & facNCCE2.RutCliente != "99501760-1" & facNCCE2.RutCliente != "99554560-8" & facNCCE2.RutCliente != "99586280-8")
                            {
                                facNCCE2.CodigoDelProducto = "110804";
                                facNCCE2.CentroDeCostos = "209";
                            }


                            facturasNCCE.Add(facNCCE2);
                            

                        }
                    }
                    

               
                    //f.MontoExento = exentoAntes;
                    

                    if (f.MontoExento !="0" && String.IsNullOrEmpty(f.MontoExento)==false && f.TipoDeDocumento == "FACE"  && esFacturaDeEnvases==false)
                    {
                        f = new Factura();
                        sFileName = item;

                        f.MontoAfecto = "0";
                        f.MontoExento = "0";
                        f.MontoIva = "0";
                        f.TotalDelDocumento = "0";

                        f.TipoDeDocumento = getValue("TipoDTE", sFileName);

                        f.TipoDeDocumento = determinarTipoDeDocumento(f.TipoDeDocumento);
                        

                        f.NumeroDelDocumento = getValue("Folio", sFileName);
                        f.FechaDeDocumento = convertirAFechaValida(getValue("FchEmis", sFileName));
                        f.FechaContableDeDocumento = convertirAFechaValida(getValue("FchEmis", sFileName));//que es la fecha de cancelacion?
                        f.FechaDeVencimientoDeDocumento = convertirAFechaValida(getValue("FchEmis", sFileName));//convertirAFechaValida(getValue("FchVenc", sFileName));// fecha de vencimiento debe ser igual o mayor a fecha de emision

                        DateTime fechaActual = DateTime.Now;
                        f.FechaContableDeDocumento = convertirAFechaValidaDesdeTranstecnia(Convert.ToString(now.Date));//"dia actual"

                        f.FechaDeDocumento = f.FechaContableDeDocumento;
                        f.FechaDeVencimientoDeDocumento = f.FechaContableDeDocumento;

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



                        if (f.TipoDeDocumento == "FCEE")
                        {
                            f.MontoExento = getValue("MntExe", sFileName);

                        }


                        f.MontoIva = getValue("IVA", sFileName);
                        f.MontoImpuestosEspecificos = "";//getValue("Folio", sFileName);
                        f.MontoIvaRetenido = "";//getValue("Folio", sFileName);
                        f.MontoImpuestosRetenidos = "";// getValue("TipoDTE", sFileName);
                        f.TipoDeDescuentoGlobal = "";//getValue("Folio", sFileName);
                        f.DescuentoGlobal = "";//getValue("Folio", sFileName);
                        f.TotalDelDocumento = getValue("MntTotal", sFileName);
                        f.DeudaPendiente = getValue("MntTotal", sFileName);
                        f.TipoDocReferencia = "";//getValue("Folio", sFileName);
                        f.NumDocReferencia = "";//getValue("Folio", sFileName);
                        f.FechaDocReferencia = "";//getValue("Folio", sFileName);
                        f.CodigoDelProducto = "420710";//getValue("TipoDTE", sFileName);
                        f.Cantidad = "1"; //getValue("Folio", sFileName);
                        f.Unidad = "S.U.M"; //getValue("Folio", sFileName);
                        f.PrecioUnitario = getValue("MntNeto", sFileName);
                        f.MonedaDelDetalle = "CLP";
                        f.TasaDeCambio2 = "1";//getValue("TipoDTE", sFileName);
                        f.NumeroDeSerie = "";//getValue("Folio", sFileName);
                        f.NumeroDeLote = "";//getValue("Folio", sFileName);
                        f.FechaDeVencimiento = "";// getValue("Folio", sFileName);
                        f.CentroDeCostos = getValue("DirOrigen", sFileName); //Este es el centro de costos


                        //si el rut del receptor es 78462150-2, el costeo es para TCP, si
                        //es 78877610-1, es para PSCP
                         rutDeReceptor = getValue("RutReceptor", sFileName);
                         esPSCP = false;
                        if (rutDeReceptor == "78877610-1")
                        {
                            esPSCP = true;
                        }

                        f.CentroDeCostos = determinarCentroDeCosto(f.CentroDeCostos, esPSCP);

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
                        f.Glosa = "Factura de compra";//getValue("Folio", sFileName);
                        f.Referencia = "";//getValue("Folio", sFileName);
                        f.FechaDeComprometida = "";//getValue("Folio", sFileName);
                        f.PorcentajeCEEC = "";//getValue("Folio", sFileName);
                        f.ImpuestoLey18211 = "";
                        f.IvaLey18211 = "";
                        f.CodigoKitFlexible = "";
                        f.AjusteIva = "";

                        f.PrecioUnitario = f.MontoExento;
                        f.CodigoDelProducto = "420724E";


                        if (f.RutCliente != "91041000-8" & f.RutCliente != "96989120-4" & f.RutCliente != "99501760-1" & f.RutCliente != "99554560-8" & f.RutCliente != "99586280-8")
                        {
                            f.CodigoDelProducto = "110804";
                            f.CentroDeCostos = "209";
                        }


                        facturas.Add(f);


                    }


              



                }


            }


            String pathDeDescargas = getCarpetaDeDescargas();
            pathDeDescargas = pathDeDescargas + "" + @"\FacturasEnExcelAPartirDeXml.xlsx";
            var archivo = new FileInfo(pathDeDescargas);
            SaveExcelFile(facturas, archivo);

            pathDeDescargas = getCarpetaDeDescargas()+ "" + @"\FacturasNCCEEnExcelAPartirDeXml.xlsx";
             archivo = new FileInfo(pathDeDescargas);
       
            SaveExcelFileNCCE(facturasNCCE, archivo);
            MessageBox.Show("Archivo FacturasEnExcelAPartirDeXml y archivo de FacturasNCCEEnExcelAPartirDeXML creado en carpeta de descargas!");


        }


    

        private static async Task SaveExcelFile(List<Factura> facturas, FileInfo file)
        {
            var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("Facturas");

            var range = ws.Cells["A1"].LoadFromCollection(facturas, true);

            range.AutoFitColumns();

            await package.SaveAsync();
        }

        private static async Task SaveExcelFileNCCE(List<FacturaNCCE> ncces, FileInfo file)
        {
            var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("Facturas");

            var range = ws.Cells["A1"].LoadFromCollection(ncces, true);

            range.AutoFitColumns();

            await package.SaveAsync();
        }

        private static async Task SaveExcelFileMatches(List<MatchFacturaManagerCTACTE> coincidencias, FileInfo file)
        {
            var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("Match");

            var range = ws.Cells["A1"].LoadFromCollection(coincidencias, true);

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

            //manager pide dd/mm/yyyy, pero en la ultima prueba solo tomo mm/dd/yyyy

            //MessageBox.Show(datos[0]);//ano
            //MessageBox.Show(datos[1]);//mes
            //MessageBox.Show(datos[2]);//dia
            fechaValida = datos[1]+"/"+datos[2]+"/" + datos[0];

            // return fechaValida;
            //lo cambie porque está raro
            return "1/28/2022";

        }

        public String convertirAFechaValida2(String fechaAConvertir)
        {
            String fechaValida = "";

            string[] datos = fechaAConvertir.Split('/');

            //manager pide dd/mm/yyyy, pero en la ultima prueba solo tomo mm/dd/yyyy

            fechaValida = datos[1] + "/" + datos[0] + "/" + datos[2];

             return fechaValida;
      

        }


        public String convertirAFechaValidaDesdeTranstecnia(String fechaAConvertir)
        {
            String fechaValida = "";

            string[] fechaDeTranstecnia = fechaAConvertir.Split(' ');

            String soloLaFechaDeTranstecnia = fechaDeTranstecnia[0];

            string [] partes= soloLaFechaDeTranstecnia.Split('/');

            fechaValida = partes[1]+"/"+ partes[0] + "/" + partes[2] + "";


             return fechaValida;
        

        }

        public String validarRutQueVieneDeTranstecnia(String rutAValidar)
        {

            if (rutAValidar.Length > 0) { rutAValidar = rutAValidar.Insert(rutAValidar.Length - 1, "-"); }

            return rutAValidar.Trim();
        }



        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Funcionalidad incompleta");

            //No estoy seguro de que se pueda manipular adecuadamente la información de un PDF

            List<Factura> facturas = new List<Factura>();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Se generará un Excel a partir de la información presente en la base de datos de Transtecnia.");

            List<Factura> facturas = new List<Factura>();
            Factura f = new Factura();

            int ano = 2022;
            //el rut de Copec es 995200007


            string str = @"Data Source=172.16.1.198\SQLEXPRESS;Initial Catalog=C001" + ano + ";User ID=sa;Password=Super123";
            Console.WriteLine(str);
            SqlConnection con = new SqlConnection(str);
            SqlCommand cmd = new SqlCommand();
            SqlDataReader r;
            
            //En transtecnia ImpCod = 3 es especifico no recuperable (69%) e ImpCod= 1 es iva (19%)
            string sql = @"SELECT 
              A.DocCod
              ,A.ECVNumDoc
              ,A.CpRut
              ,B.CtaCod
              ,B.CtroCod
              ,B.DCVGlosa
              ,C.ECVFecha
              ,C.ECVVence
              ,C.ECVExento
              ,C.ECVNeto
              ,A.ICVMonto
              ,A.ICVBase
              ,A.ImpCod
               FROM [C0012022].[dbo].[ImpCpaVta]AS A
               JOIN [C0012022].[dbo].[DetCpaVta] AS B
               ON A.ECVNumDoc=B.ECVNumDoc AND A.CpRut=B.CpRut
               JOIN [C0012022].[dbo].[EncCpaVta] AS C
               ON A.ECVNumDoc=C.ECVNumDoc AND A.CpRut=C.CpRut
                AND A.CpRut= ' 995200007'
                ";




            //string sql = @"SELECT 
            //  A.DocCod
            //  ,A.ECVNumDoc
            //  ,A.CpRut
            //  ,B.CtaCod
            //  ,B.CtroCod
            //  ,B.DCVGlosa
            //  ,C.ECVFecha
            //  ,C.ECVVence
            //  ,C.ECVExento
            //  ,C.ECVNeto
            //  ,A.ICVMonto
            //   FROM [C0012022].[dbo].[ImpCpaVta]AS A
            //   JOIN [C0012022].[dbo].[DetCpaVta] AS B
            //   ON A.ECVNumDoc=B.ECVNumDoc AND A.CpRut=B.CpRut
            //   JOIN [C0012022].[dbo].[EncCpaVta] AS C
            //   ON A.ECVNumDoc=C.ECVNumDoc AND A.CpRut=C.CpRut
            //    AND A.CpRut= ' 995200007'
            //    ";

            string sql2 = @"SELECT 
       A.DocCod
      ,A.ECVNumDoc
      ,A.CpRut
      ,A.CtaCod
      ,A.DCVMonto
      ,A.CtroCod
      ,A.DCVTri
      ,A.DCVActF
      ,A.DCVGlosa
      ,B.ECVFecha
      ,B.ECVExento
      ,B.ECVNeto

  FROM [C0012022].[dbo].[DetCpaVta] AS A
   JOIN [C0012022].[dbo].[EncCpaVta] AS B
   ON  A.ECVNumDoc=B.ECVNumDoc AND A.CpRut=B.CpRut
   WHERE A.DocCod=12
    AND A.CpRut= ' 995200007'";

            //    A.DocCod=12        AND(A.ECVNumDoc = 156 OR

            //A.ECVNumDoc = 268 OR

            //A.ECVNumDoc = 265 OR

            //A.ECVNumDoc = 9555467 OR

            //A.ECVNumDoc = 9561089 OR

            //A.ECVNumDoc = 9561250 OR

            //A.ECVNumDoc = 7363614 OR

            //A.ECVNumDoc = 7366342 OR

            //A.ECVNumDoc = 4324925 OR

            //A.ECVNumDoc = 104401 OR

            //A.ECVNumDoc = 173 OR

            //A.ECVNumDoc = 1171
            // )



            Console.WriteLine("[" + sql + "]");
            Console.WriteLine("[" + sql2 + "]");

            cmd.CommandText = sql;
            cmd.CommandType = CommandType.Text;
            cmd.Connection = con;
            try
            {
                con.Open();
            }
            catch (Exception)
            {
                MessageBox.Show("Problema de comunicacion con el Servidor.  Por favor revise su conexion a Internet o VPN.");
            }

            r = cmd.ExecuteReader();
            while (r.Read())
            {
                f = new Factura();

                f.MontoAfecto = "0";
                f.MontoExento = "0";
                f.MontoIva = "0";
                f.TotalDelDocumento = "0";



                f.TipoDeDocumento = Convert.ToString(r.GetValue(0));//depende del numero  DocCod
                f.TipoDeDocumento=f.TipoDeDocumento.Trim();


                f.TipoDeDocumento = determinarTipoDeDocumentoProvenienteDeTranstecnia(f.TipoDeDocumento);

                f.NumeroDelDocumento = Convert.ToString(r.GetValue(1));

                //Ojo con estas fechas

                f.FechaDeDocumento = convertirAFechaValidaDesdeTranstecnia(Convert.ToString(r.GetValue(6)));
                DateTime now = DateTime.Now;
                f.FechaContableDeDocumento =   convertirAFechaValidaDesdeTranstecnia(Convert.ToString(now.Date));//"dia actual"
                f.FechaDeVencimientoDeDocumento = convertirAFechaValidaDesdeTranstecnia(Convert.ToString(r.GetValue(7)));

                f.FechaDeDocumento = f.FechaContableDeDocumento;
                f.FechaDeVencimientoDeDocumento = f.FechaContableDeDocumento;

                //que es la unidad de negocio?
                f.CodigoDeUnidadDeNegocio = "1";
                f.RutCliente = validarRutQueVieneDeTranstecnia(Convert.ToString(r.GetValue(2)));
                f.DireccionDelCliente = "Casa Matriz"; 
                f.RutFacturador = "";
                f.CodigoVendedor = "";
                f.CodigoComisionista = "";
                f.Probabilidad = "";
                f.ListaPrecio = "";
                f.PlazoPago = "P01";
                f.MonedaDelDocumento = "CLP";
                f.TasaDeCambio = "";
                //hay que alterar cosas aqui
                f.MontoAfecto = Convert.ToString(r.GetValue(9));//ECVNeto

                f.MontoExento = Convert.ToString(r.GetValue(8));//ECVExento
                f.MontoImpuestosEspecificos = "";

                String codigoDeImpuesto = Convert.ToString(r.GetValue(12)); // ImpCod


                //410104   Petroleo Gasto
                //110904  Impuesto Específico Generico

                if (codigoDeImpuesto == "1")
                {
                    f.MontoIva = Convert.ToString(r.GetValue(10));//ICVMonto
                    f.CodigoDelProducto = "410104";
                }
                else if(codigoDeImpuesto=="3")
                {
                    f.MontoIva = "0";
                    f.MontoImpuestosEspecificos = Convert.ToString(r.GetValue(10));
                    f.CodigoDelProducto = "110904";
                }

                //parece que el codigo de producto especifico no esta asociado al producto???

                // Convert.ToString(r.GetValue(11)); // ICVBase



                f.MontoIvaRetenido = "";
                f.MontoImpuestosRetenidos = "";
                f.TipoDeDescuentoGlobal = "";
                f.DescuentoGlobal = "";
                f.TotalDelDocumento = Convert.ToString(Convert.ToInt32(f.MontoAfecto) + Convert.ToInt32(f.MontoExento) + Convert.ToInt32(f.MontoIva));  //afecto (o neto) + exento + iva
                f.DeudaPendiente = f.TotalDelDocumento; //esto es el monto total
                f.TipoDocReferencia = "";
                f.NumDocReferencia = "";
                f.FechaDocReferencia = "";

                f.CodigoDelProducto = "420710";

                //if (f.TipoDeDocumento == "FACE" ^ f.TipoDeDocumento == "NCCE")
                //{
                //    f.CodigoDelProducto = "420710";
               

                //}else if (f.TipoDeDocumento == "FCEE")
                //{
                //    f.CodigoDelProducto = "420724E";
                //}
                
                
                f.Cantidad = "1"; 
                f.Unidad = "S.U.M"; 
                f.PrecioUnitario = f.MontoAfecto;
                f.MonedaDelDetalle = "CLP";
                f.TasaDeCambio2 = "1";
                f.NumeroDeSerie = "";
                f.NumeroDeLote = "";
                f.FechaDeVencimiento = "";
                f.CentroDeCostos = Convert.ToString(r.GetValue(4));//sacar de transtecnia, es un numero CtroCod


                String rutDeReceptor = "78462150-2";
                Boolean esPSCP = false;
                if (rutDeReceptor == "78877610-1")
                {
                    esPSCP = true;
                }

                f.CentroDeCostos = determinarCentroDeCostoProvenienteDeTranstecnia(f.CentroDeCostos, esPSCP);

                f.TipoDeDescuento = "";
                f.Descuento = "";
                f.Ubicacion = "";
                f.Bodega = "";
                f.Concepto1 = "";
                f.Concepto2 = "";
                f.Concepto3 = "";
                f.Concepto4 = "";
                f.Descripcion = "";
                f.DescripcionAdicional = "";
                f.Stock = "0";
                f.Comentario11 = "";
                f.Comentario21 = "";
                f.Comentario31 = "";
                f.Comentario41 = "";
                f.Comentario51 = "";
                f.CodigoImpuestoEspecifico1 = "";
                f.MontoImpuestoEspecifico1 = "";
                f.CodigoImpuestoEspecifico2 = "";
                f.MontoImpuestoEspecifico2 = "";

                //Modalidad es necesaria para la nota de credito
                f.Modalidad = "";


                f.Glosa = Convert.ToString(r.GetValue(5));//sacar de transtecnia
                f.Referencia = "";
                f.FechaDeComprometida = "";
                f.PorcentajeCEEC = "";
                f.ImpuestoLey18211 = "";
                f.IvaLey18211 = "";
                f.CodigoKitFlexible = "";
                f.AjusteIva = "";


                if (Convert.ToInt32(f.MontoExento) > 0 && f.TipoDeDocumento == "FACE")
                {
                    Factura f2 = new Factura();


                    f2.TipoDeDocumento = f.TipoDeDocumento;
                    f2.NumeroDelDocumento = f.NumeroDelDocumento;
                    f2.FechaContableDeDocumento = f.FechaContableDeDocumento;       

                    f2.FechaDeDocumento = f2.FechaContableDeDocumento;
                    f2.FechaDeVencimientoDeDocumento = f2.FechaContableDeDocumento;

                    f2.CodigoDeUnidadDeNegocio = "1";
                    f2.RutCliente = f.RutCliente;
                    f2.DireccionDelCliente = "Casa Matriz";
                    f2.RutFacturador = "";
                    f2.CodigoVendedor = "";
                    f2.CodigoComisionista = "";
                    f2.Probabilidad = "";
                    f2.ListaPrecio = "";
                    f2.PlazoPago = "P01";
                    f2.MonedaDelDocumento = "CLP";
                    f2.TasaDeCambio = "";
                    f2.MontoAfecto = f.MontoAfecto;
                    f2.MontoExento = f.MontoExento;
                    f2.MontoIva = f.MontoIva;
                    f2.MontoImpuestosEspecificos = "";
                    f2.MontoIvaRetenido = "";
                    f2.MontoImpuestosRetenidos = "";
                    f2.TipoDeDescuentoGlobal = "";
                    f2.DescuentoGlobal = "";
                    f2.TotalDelDocumento = Convert.ToString(Convert.ToInt32(f2.MontoAfecto) + Convert.ToInt32(f2.MontoExento) + Convert.ToInt32(f2.MontoIva));  //afecto (o neto) + exento + iva
                    f2.DeudaPendiente = f.TotalDelDocumento;
                    f2.TipoDocReferencia = "";
                    f2.NumDocReferencia = "";
                    f2.FechaDocReferencia = "";
                    f2.Cantidad = "1";
                    f2.Unidad = "S.U.M";
                    f2.MonedaDelDetalle = "CLP";
                    f2.TasaDeCambio2 = "1";
                    f2.NumeroDeSerie = "";
                    f2.NumeroDeLote = "";
                    f2.FechaDeVencimiento = "";
                    f2.CentroDeCostos = "";
                    f2.TipoDeDescuento = "";
                    f2.Descuento = "";
                    f2.Ubicacion = "";
                    f2.Bodega = "";
                    f2.Concepto1 = "";
                    f2.Concepto2 = "";
                    f2.Concepto3 = "";
                    f2.Concepto4 = "";
                    f2.Descripcion = "";
                    f2.DescripcionAdicional = "";
                    f2.Stock = "0";
                    f2.Comentario11 = "";
                    f2.Comentario21 = "";
                    f2.Comentario31 = "";
                    f2.Comentario41 = "";
                    f2.Comentario51 = "";
                    f2.CodigoImpuestoEspecifico1 = "";
                    f2.MontoImpuestoEspecifico1 = "";
                    f2.CodigoImpuestoEspecifico2 = "";
                    f2.MontoImpuestoEspecifico2 = "";
                    f2.Modalidad = "";
                    f2.Glosa = f.Glosa;
                    f2.Referencia = "";
                    f2.FechaDeComprometida = "";
                    f2.PorcentajeCEEC = "";
                    f2.ImpuestoLey18211 = "";
                    f2.IvaLey18211 = "";
                    f2.CodigoKitFlexible = "";
                    f2.AjusteIva = "";
                    f2.CentroDeCostos = f.CentroDeCostos;
                    

                    f2.PrecioUnitario = f.MontoExento;
                    f2.CodigoDelProducto = "420724E";

                    
                    facturas.Add(f2);

                }


                facturas.Add(f);

            }
            con.Close();

            //esto es para facturas exentas

            cmd.CommandText = sql2;
            cmd.CommandType = CommandType.Text;
            cmd.Connection = con;
            try
            {
                con.Open();
            }
            catch (Exception)
            {
                MessageBox.Show("Problema de comunicacion con el Servidor.  Por favor revise su conexion a Internet o VPN.");
            }

            r = cmd.ExecuteReader();
            while (r.Read())
            {
                f = new Factura();

                f.MontoAfecto = "0";
                f.MontoExento = "0";
                f.MontoIva = "0";
                f.TotalDelDocumento = "0";



                f.TipoDeDocumento = Convert.ToString(r.GetValue(0));//depende del numero  DocCod
                f.TipoDeDocumento = f.TipoDeDocumento.Trim();


                f.TipoDeDocumento = determinarTipoDeDocumentoProvenienteDeTranstecnia(f.TipoDeDocumento);
        

                f.NumeroDelDocumento = Convert.ToString(r.GetValue(1));

                //Ojo con estas fechas

                f.FechaDeDocumento = convertirAFechaValidaDesdeTranstecnia(Convert.ToString(r.GetValue(9)));
                DateTime now = DateTime.Now;
                f.FechaContableDeDocumento = convertirAFechaValidaDesdeTranstecnia(Convert.ToString(now.Date));//"dia actual"
                f.FechaDeVencimientoDeDocumento = convertirAFechaValidaDesdeTranstecnia(Convert.ToString(r.GetValue(9)));

                f.FechaDeDocumento =f.FechaContableDeDocumento;
                f.FechaDeVencimientoDeDocumento =f.FechaContableDeDocumento;

                //que es la unidad de negocio?
                f.CodigoDeUnidadDeNegocio = "1";
                f.RutCliente = validarRutQueVieneDeTranstecnia(Convert.ToString(r.GetValue(2)));
                f.DireccionDelCliente = "Casa Matriz";
                f.RutFacturador = "";
                f.CodigoVendedor = "";
                f.CodigoComisionista = "";
                f.Probabilidad = "";
                f.ListaPrecio = "";
                f.PlazoPago = "P01";
                f.MonedaDelDocumento = "CLP";
                f.TasaDeCambio = "";
                f.MontoAfecto = Convert.ToString(r.GetValue(11));//ECVNeto

                f.MontoExento = Convert.ToString(r.GetValue(10));//ECVExento


                f.MontoIva ="0";//ICVMonto

                f.MontoImpuestosEspecificos = "";
                f.MontoIvaRetenido = "";
                f.MontoImpuestosRetenidos = "";
                f.TipoDeDescuentoGlobal = "";
                f.DescuentoGlobal = "";
                f.TotalDelDocumento = Convert.ToString(Convert.ToInt32(f.MontoAfecto) + Convert.ToInt32(f.MontoExento) + Convert.ToInt32(f.MontoIva));  //afecto (o neto) + exento + iva
                f.DeudaPendiente = f.TotalDelDocumento; //esto es el monto total
                f.TipoDocReferencia = "";
                f.NumDocReferencia = "";
                f.FechaDocReferencia = "";


                if (f.TipoDeDocumento == "FACE" ^ f.TipoDeDocumento == "NCCE")
                {
                    f.CodigoDelProducto = "420710";

                }
                else if (f.TipoDeDocumento == "FCEE")
                {
                    f.CodigoDelProducto = "420724E";
                }


                f.Cantidad = "1";
                f.Unidad = "S.U.M";
                f.PrecioUnitario = f.MontoAfecto;
                f.MonedaDelDetalle = "CLP";
                f.TasaDeCambio2 = "1";
                f.NumeroDeSerie = "";
                f.NumeroDeLote = "";
                f.FechaDeVencimiento = "";
                f.CentroDeCostos = Convert.ToString(r.GetValue(5));//sacar de transtecnia, es un numero CtroCod


                String rutDeReceptor = "78462150-2";
                Boolean esPSCP = false;
                if (rutDeReceptor == "78877610-1")
                {
                    esPSCP = true;
                }

                f.CentroDeCostos =determinarCentroDeCostoProvenienteDeTranstecnia(f.CentroDeCostos,esPSCP);

                f.TipoDeDescuento = "";
                f.Descuento = "";
                f.Ubicacion = "";
                f.Bodega = "";
                f.Concepto1 = "";
                f.Concepto2 = "";
                f.Concepto3 = "";
                f.Concepto4 = "";
                f.Descripcion = "";
                f.DescripcionAdicional = "";
                f.Stock = "0";
                f.Comentario11 = "";
                f.Comentario21 = "";
                f.Comentario31 = "";
                f.Comentario41 = "";
                f.Comentario51 = "";
                f.CodigoImpuestoEspecifico1 = "";
                f.MontoImpuestoEspecifico1 = "";
                f.CodigoImpuestoEspecifico2 = "";
                f.MontoImpuestoEspecifico2 = "";

                //Modalidad es necesaria para la nota de credito
                f.Modalidad = "";


                f.Glosa = Convert.ToString(r.GetValue(8));//sacar de transtecnia
                f.Referencia = "";
                f.FechaDeComprometida = "";
                f.PorcentajeCEEC = "";
                f.ImpuestoLey18211 = "";
                f.IvaLey18211 = "";
                f.CodigoKitFlexible = "";
                f.AjusteIva = "";

                

                //if (Convert.ToInt32(f.MontoExento) > 0)
                //{
                //    Factura f2 = new Factura();
                //    f2 = f;
                //    f2.PrecioUnitario = f.MontoExento;
                //    f.CodigoDelProducto = "420710";

                //    f2.CodigoDelProducto = "420724E";
                //    facturas.Add(f2);

                //}

                facturas.Add(f);


            }
            con.Close();


            //termino de seccion para facturas exentas



            String pathDeDescargas = getCarpetaDeDescargas();
            pathDeDescargas = pathDeDescargas + "" + @"\FacturasEnExcelAPartirDeTranstecnia.xlsx";
            var archivo = new FileInfo(pathDeDescargas);
            SaveExcelFile(facturas, archivo);
            MessageBox.Show("Archivo FacturasEnExcelAPartirDeTranstecnia creado en carpeta de descargas!");


        }

        private void button4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Funcionalidad incompleta");




            
        }



        private String determinarCentroDeCostoProvenienteDeTranstecnia(String centroDeCostoComoNumeroQueEsString, Boolean esPSCP)
        {
            String centroDeCosto = "";


            switch (centroDeCostoComoNumeroQueEsString)
            {
                case "5":
                    centroDeCosto = "202";
                    if (esPSCP)
                    {
                        centroDeCosto = "302";
                    }
                    break;
                case "3":
                    centroDeCosto = "201";
                    if (esPSCP)
                    {
                        centroDeCosto = "302";
                    }
                    break;
                case "11":
                    centroDeCosto = "200";
                    if (esPSCP)
                    {
                        centroDeCosto = "300";
                    }
                    break;
                case "10":
                    centroDeCosto = "207";
                    if (esPSCP)
                    {
                        centroDeCosto = "302";
                    }
                    break;
                case "16"://santiago sur es eccusa?
                    centroDeCosto = "206";
                    if (esPSCP)
                    {
                        centroDeCosto = "302";
                    }
                    break;
                case "9":
                    //Santiago es emprendedor?
                    centroDeCosto = "206";
                    if (esPSCP)
                    {
                        centroDeCosto = "306";
                    }
                    break;
                case "7":
                    centroDeCosto = "205";
                    if (esPSCP)
                    {
                        centroDeCosto = "305";
                    }
                    break;
                case "1":
                    centroDeCosto = "204";
                    if (esPSCP)
                    {
                        centroDeCosto = "304";
                    }
                    break;
                case "12":
                    centroDeCosto = "204";
                    if (esPSCP)
                    {
                        centroDeCosto = "304";
                    }
                    break;
                case "13":
                    centroDeCosto = "204";
                    if (esPSCP)
                    {
                        centroDeCosto = "304";
                    }
                    break;
                case "6":
                    centroDeCosto = "204";
                    if (esPSCP)
                    {
                        centroDeCosto = "304";
                    }
                    break;
                default:
                    centroDeCosto = "204";
                    if (esPSCP)
                    {
                        centroDeCosto = "304";
                    }
                    break;

            }



            return centroDeCosto;


        }


        private String determinarCentroDeCosto(String centroDeCostoComoString, Boolean esPSCP)
        {
            String centroDeCosto = "";
            //se supone que a veces vienen cosas como STGO PN8000


            switch (centroDeCostoComoString)
            {
                case "CURICO":
                    centroDeCosto = "202";
                    if (esPSCP)
                    {
                        centroDeCosto = "302";
                    }
                    break;
                case "RANCAGUA":
                    centroDeCosto = "201";
                    if (esPSCP)
                    {
                        centroDeCosto = "302";
                    }
                    break;
                case "MELIPILLA":
                    centroDeCosto = "200";
                    if (esPSCP)
                    {
                        centroDeCosto = "300";
                    }
                    break;
                case "SAN ANTONIO":
                    centroDeCosto = "207";
                    if (esPSCP)
                    {
                        centroDeCosto = "307";
                    }
                    break;
                case "SANTIAGO SUR":
                    centroDeCosto = "206";
                    if (esPSCP)
                    {
                        centroDeCosto = "306";
                    }
                    break;
                case "SANTIAGO":
                    //Santiago es emprendedor?
                    centroDeCosto = "206";
                    if (esPSCP)
                    {
                        centroDeCosto = "306";
                    }
                    break;
                case "ILLAPEL":
                    centroDeCosto = "205";
                    if (esPSCP)
                    {
                        centroDeCosto = "305";
                    }
                    break;
                default:
                    //Interplantas, cuando no es ninguno de los anteriores
                    centroDeCosto = "204";//"204";
                    if (esPSCP)
                    {
                        centroDeCosto = "304";//"304";
                    }
                    break;

            }

            return centroDeCosto;

        }


        private String determinarTipoDeDocumento(String codigoDeDocumento)
        {
            String tipoDeFactura= "";


            switch (codigoDeDocumento)
            {
                case "33":
                    tipoDeFactura = "FACE";
                    break;
                case "34":
                    tipoDeFactura = "FCEE";
                    break;
                case "61":
                    tipoDeFactura = "NCCE";
                    break;
                case "52":
                    //es guia de despacho
                    tipoDeFactura = "guia de despacho";
                    break;
                default:
                    tipoDeFactura = codigoDeDocumento;
                    break;
            }


            return tipoDeFactura;


        }



        private String determinarTipoDeDocumentoProvenienteDeTranstecnia(String codigoDeDocumento)
        {
            String tipoDeFactura = "";


            switch (codigoDeDocumento)
            {
                case "4":
                    tipoDeFactura = "FACE";
                    break;
                case "12":
                    tipoDeFactura = "FCEE";
                    break;
                case "18":
                    tipoDeFactura = "NCCE";
                    break;
                default:
                    break;
            }


            return tipoDeFactura;


        }


        private String getCarpetaDeDescargas()
        {
            string localfolder = ApplicationData.Current.LocalFolder.Path;
            var array = localfolder.Split('\\');
            var username = array[2];
            string downloads = @"C:\Users\" + username + @"\Downloads";
          

            return downloads;
        }

        private FacturaNCCE convertirFacturaANCCE(Factura f)
        {
            FacturaNCCE facNCCE = new FacturaNCCE();

            facNCCE.TipoDeDocumento = f.TipoDeDocumento;
            facNCCE.NumeroDelDocumento = f.NumeroDelDocumento;
            facNCCE.FechaDeDocumento = f.FechaDeDocumento;
            facNCCE.FechaDeContableDeDocumento = f.FechaContableDeDocumento;
            facNCCE.FechaDeVencimientoDeDocumento = f.FechaDeVencimientoDeDocumento;
            facNCCE.CodigoUnidadDeNegocio = f.CodigoDeUnidadDeNegocio;
            facNCCE.RutCliente = f.RutCliente;
            facNCCE.DireccionCliente = f.DireccionDelCliente;
            facNCCE.RutFacturador = f.RutFacturador;
            facNCCE.CodigoVendedor = f.CodigoVendedor;
            facNCCE.CodigoComisionista = f.CodigoComisionista;
            facNCCE.Probablidad = f.Probabilidad;
            facNCCE.ListaPrecio = f.ListaPrecio;
            facNCCE.PlazoPago = f.PlazoPago;
            facNCCE.MonedaDelDocumento = f.MonedaDelDocumento;
            facNCCE.TasaDeCambio = f.TasaDeCambio;
            facNCCE.MontoAfecto = f.MontoAfecto;
            facNCCE.MontoExento = f.MontoExento;
            facNCCE.MontoIva = f.MontoIva;
            facNCCE.MontoImpuestosEspecificos = f.MontoImpuestosEspecificos;
            facNCCE.MontoIvaRetenido = f.MontoIvaRetenido;
            facNCCE.MontoImpuestosRetenidos = f.MontoImpuestosRetenidos;
            facNCCE.TipoDeDescuentoGlobal = f.TipoDeDescuentoGlobal;
            facNCCE.DescuentoGlobal = f.DescuentoGlobal;
            facNCCE.TotalDelDocumento = f.TotalDelDocumento;
            facNCCE.DeudaPendiente = f.DeudaPendiente;
            facNCCE.TipoDocumentoReferencia = f.TipoDocReferencia;
            facNCCE.NumDocReferencia = f.NumDocReferencia;
            facNCCE.FechaDocumentoDeReferencia = f.FechaDocReferencia;
            facNCCE.CodigoDelProducto = f.CodigoDelProducto;
            facNCCE.Cantidad = f.Cantidad;
            facNCCE.Unidad = f.Unidad;
            facNCCE.PrecioUnitario = f.PrecioUnitario;
            facNCCE.MonedaDelDetalle = f.MonedaDelDetalle;
            facNCCE.TasaDeCambio2 = f.TasaDeCambio2;
            facNCCE.NumeroDeSerie = f.NumeroDeSerie;
            facNCCE.NumeroDeLote = f.NumeroDeLote;
            facNCCE.FechaDeVencimiento = f.FechaDeVencimiento;
            facNCCE.CentroDeCostos = f.CentroDeCostos;
            facNCCE.TipoDeDescuento = f.TipoDeDescuento;
            facNCCE.Descuento = f.Descuento;
            facNCCE.Ubicacion = f.Ubicacion;
            facNCCE.Bodega = f.Bodega;
            facNCCE.Concepto1 = f.Concepto1;
            facNCCE.Concepto2 = f.Concepto2;
            facNCCE.Concepto3 = f.Concepto3;
            facNCCE.Concepto4 = f.Concepto4;
            facNCCE.Descripcion = f.Descripcion;
            facNCCE.DescripcionAdicional = f.DescripcionAdicional;
            facNCCE.Stock = f.Stock;
            facNCCE.Comentario1 = f.Comentario11;
            facNCCE.Comentario2 = f.Comentario21;
            facNCCE.Comentario3 = f.Comentario31;
            facNCCE.Comentario4 = f.Comentario41;
            facNCCE.Comentario5 = f.Comentario51;
            facNCCE.CodigoImpEspecial1 = f.CodigoImpuestoEspecifico1;
            facNCCE.MontoImpEspecial1 = f.MontoImpuestoEspecifico1;
            facNCCE.CodigoImpEspecial2 = f.CodigoImpuestoEspecifico2;
            facNCCE.MontoImpEspecial2 = f.MontoImpuestoEspecifico2;
            facNCCE.Modalidad = f.Modalidad;
            facNCCE.Glosa = f.Glosa;
            facNCCE.Referencia = f.Referencia;
            facNCCE.FechaDeComprometida = f.FechaDeComprometida;
            facNCCE.PorcentajeCEEC = "";
            facNCCE.TipoDeDocumentoDeOrigen = "";
            facNCCE.NumeroDocumentoDeOrigen = "";
            facNCCE.NumeroDetalleOrigen = "";
            facNCCE.CodigoKitFlexible = f.CodigoKitFlexible;
            facNCCE.AjusteIva = f.AjusteIva;

            return facNCCE;
        }

        private void button5_Click(object sender, EventArgs e)
        {

            List<FacturaManager> facturasDeManager = new List<FacturaManager>();
            List<CobroCTACTE> cobrosCTACTE = new List<CobroCTACTE>();

            List<MatchFacturaManagerCTACTE> listadoDeFacturasParaElTercerExcel = new List<MatchFacturaManagerCTACTE>();

            string sFileName = "";

            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "Archivos XLSX (*.xlsx)|*.xlsx";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            string[] arrAllFiles = new string[] { };

            MessageBox.Show("Seleccionar excel de facturas de manager");
            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                sFileName = choofdlog.FileName;
                arrAllFiles = choofdlog.FileNames; //used when Multiselect = true           
            }


            facturasDeManager = leerExcelDeFacturasDeManager(sFileName);

            MessageBox.Show("Seleccionar excel de cobros");
            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                sFileName = choofdlog.FileName;
                arrAllFiles = choofdlog.FileNames; //used when Multiselect = true           
            }
            cobrosCTACTE = leerExcelDeCobrosCTACTE(sFileName);

            
            foreach (var item in facturasDeManager)
            {

                foreach (var i in cobrosCTACTE)
                {
                    if((item.Total==i.Monto) && (item.FechaDoc1==i.Fecha))
                    {
                        MatchFacturaManagerCTACTE matchFacturaManagerCTACTE = new MatchFacturaManagerCTACTE();

                        matchFacturaManagerCTACTE.Fecha = i.Fecha;
                        matchFacturaManagerCTACTE.Monto = i.Monto;
                        matchFacturaManagerCTACTE.Centro = i.CtoCosto;

                        switch (matchFacturaManagerCTACTE.Centro)
                        {
                            case "CURICO":
                                matchFacturaManagerCTACTE.Centro ="202";
                                break;
                            case "RANCAGUA":
                                matchFacturaManagerCTACTE.Centro = "201";
                                break;
                            case "MELIPILLA":
                                matchFacturaManagerCTACTE.Centro = "200";
                                break;
                            case "SAN ANTONIO":
                                matchFacturaManagerCTACTE.Centro = "207";
                                break;
                            case "SANTIAGO":
                                matchFacturaManagerCTACTE.Centro = "206";
                                break;
                            case "SANTIAGO-SUR":
                                matchFacturaManagerCTACTE.Centro = "206";
                                break;
                            case "ILLAPEL":
                                matchFacturaManagerCTACTE.Centro = "205";
                                break;
                            default:
                                //administracion
                                matchFacturaManagerCTACTE.Centro = "204";
                                break;
      
                        }

                        listadoDeFacturasParaElTercerExcel.Add(matchFacturaManagerCTACTE);
                    }
                }




            }



            //aqui habria que guardar el Excel

            String pathDeDescargas = getCarpetaDeDescargas();
            pathDeDescargas = pathDeDescargas + "" + @"\Match de facturas.xlsx";
            var archivo = new FileInfo(pathDeDescargas);
            SaveExcelFileMatches(listadoDeFacturasParaElTercerExcel, archivo);
            MessageBox.Show("Archivo Match de facturas entre Manager y CTA cte creado en carpeta de descargas!");



        }


        private List<CobroCTACTE> leerExcelDeCobrosCTACTE(String filePath)
        {
            List<CobroCTACTE>listadoDeCobros=new List<CobroCTACTE>();

            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                CobroCTACTE  encabezado = new CobroCTACTE();
                encabezado.Fecha = "Fecha";
                encabezado.CtoCosto = "CtoCosto";
                encabezado.Monto = "Monto";

                listadoDeCobros.Add(encabezado);


                for (int row = 1; row <= rowCount; row++)
                {

                    CobroCTACTE c = new CobroCTACTE();
                    c.Fecha = worksheet.Cells[row, 1].Value?.ToString().Trim();
                    c.CtoCosto = worksheet.Cells[row, 2].Value?.ToString().Trim();
                    c.Monto = worksheet.Cells[row, 3].Value?.ToString().Trim();
               
                    if (c.Monto != "Monto")
                    {
                        listadoDeCobros.Add(c);
                    }

                }
            }



            return listadoDeCobros;

        }


        private List<FacturaManager> leerExcelDeFacturasDeManager(String filePath)
        {
            List<FacturaManager> listadoDeFacturasDeManager = new List<FacturaManager>();

            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                FacturaManager encabezado = new FacturaManager();
                encabezado.Numero = "Numero";
                encabezado.Rut = "Rut";
                encabezado.Proveedor = "Proveedor";
                encabezado.Moneda = "Moneda";
                encabezado.TipoCambio = "TipoCambio";
                encabezado.FechaDoc1 = "Fecha";
                encabezado.Total = "Total";
                encabezado.Estado = "Estado";
                encabezado.Glosa = "Glosa";



                listadoDeFacturasDeManager.Add(encabezado);


                for (int row = 1; row <= rowCount; row++)
                {

                    FacturaManager fm = new FacturaManager();
                    fm.Numero = worksheet.Cells[row, 1].Value?.ToString().Trim();
                    fm.FechaDoc1 = worksheet.Cells[row, 6].Value?.ToString().Trim();
                    fm.Total = worksheet.Cells[row, 7].Value?.ToString().Trim();

                    if (fm.Total != "Total")
                    {
                        listadoDeFacturasDeManager.Add(fm);
                    }

                }
            }



            return listadoDeFacturasDeManager;

        }

    }




}
