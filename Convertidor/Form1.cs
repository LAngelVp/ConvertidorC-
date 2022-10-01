using System.Reflection;
using System.Xml;
using System.Xml.Linq;
using System;
using System.Reflection.Metadata;
using SpreadsheetLight;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Convertidor
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //Ingresamos el xml por medio de openfiledialog
        public OpenFileDialog open = new OpenFileDialog();
        public string ruta="";
        private void button1_Click(object sender, EventArgs e)
        {
            //Intentamos abrir el documento
            try
            {
                open.InitialDirectory = "C://";
                open.Filter = "Archivos XML (*.xml) |*.xml"; //Restringimos, tipos de archivos.
                if (open.ShowDialog() == DialogResult.OK)
                {
                    ruta = open.FileName;
                }
                else
                {
                    MessageBox.Show("Para realizar la conversión debe de seleccionar un archivo con extensión .xml", "Error de Carga", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception)
            {

                MessageBox.Show("Hubo un error en el proceso. Intentalo de nuevo", "Error de Carga de Archivo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
                XmlDocument xmlDoc = new();
                SLDocument documento = new SLDocument();
                SLStyle style = documento.CreateStyle();
                style.Alignment.Horizontal = HorizontalAlignmentValues.Center;
                style.Alignment.Vertical = VerticalAlignmentValues.Center;
                System.Data.DataTable tabla = new System.Data.DataTable();
                XmlReader Archivo = XmlReader.Create(ruta);
                //Creación de la Cabecera del documento.
                documento.SetCellValue("B1", "KENWORTH DEL ESTE S.A DE C.V");
                string Date = DateTime.Now.ToString("dd-MM-yyyy");
                documento.SetCellValue("D1", "FECHA");
                documento.SetCellValue("D2", Date);
                documento.MergeWorksheetCells("B1", "C2");
                //Creación de los cabeceros de la tabla.
                tabla.Columns.Add("Cantidad", typeof(string));
                tabla.Columns.Add("Descripcion", typeof(string));
                tabla.Columns.Add("No. Parte", typeof(string));
                tabla.Columns.Add("Total", typeof(string));
                //Asignación de los taños de las columnas
                documento.SetColumnWidth(2, 60);
                documento.SetColumnWidth(3, 25);
                documento.SetColumnWidth(4, 20);
                //Variables
                string ClaveProdServ = "";
                string Cantidad = "";
                string NoIdentificacion = "";
                string ClaveUnidad = "";
                string Unidad = "";
                string Descripcion = "";
                string ValorUnitario = "";
                double Importe = 0.0;
                double Total = 0.0;
                double Descuento = 0.0;
                double iva = 0.0;
                double totalGeneral = 0.0;
                double subTotal = 0.0;
            //Intento de la funcionalidad del sistema.
            try
            {
                while (Archivo.Read()) //Mientras lea el archivo
                {
                    if ((Archivo.NodeType == XmlNodeType.Element) && (Archivo.Name == "cfdi:Comprobante")) // Y tenga el nodo cfdi
                    {
                        if (Archivo.HasAttributes) //Y el nodo tenga atributos
                        {
                            Total = Convert.ToDouble(Archivo.GetAttribute("Total")); //Asignamos el total del documento a la variable
                        }
                    }
                    if ((Archivo.NodeType == XmlNodeType.Element) && (Archivo.Name == "cfdi:Concepto")) //Si el nodo cfdi:concepto tiene Elementos
                    {
                        if (Archivo.HasAttributes)//Y este tiene atributos
                        {
                            //Asignamos a variables cada atributo del elemento.
                            ClaveProdServ = Archivo.GetAttribute("ClaveProdServ");
                            NoIdentificacion = Archivo.GetAttribute("NoIdentificacion");
                            Cantidad = Archivo.GetAttribute("Cantidad");
                            ClaveUnidad = Archivo.GetAttribute("ClaveUnidad");
                            Unidad = Archivo.GetAttribute("Unidad");
                            Descripcion = Archivo.GetAttribute("Descripcion");
                            ValorUnitario = Archivo.GetAttribute("ValorUnitario");
                            Importe = Convert.ToDouble(Archivo.GetAttribute("Importe"));
                            Descuento = Convert.ToDouble(Archivo.GetAttribute("Descuento"));
                            //Operaciones
                            if (Total >= Convert.ToDouble(10000)) //Si total es mayor a Diez Mil
                            {
                                if (Descuento != Convert.ToDouble(0.0))
                                {
                                    Importe = Importe - Descuento;
                                }
                                double importeMayor = Importe / 0.9;
                                tabla.Rows.Add(Cantidad, Descripcion, NoIdentificacion, importeMayor); //Introducimos al Excel las variables.
                                subTotal += importeMayor; //Suma de la misma variable con importe mayor.
                            }
                            else
                            {
                                if (Descuento != Convert.ToDouble(0.0))
                                {
                                    Importe = Importe - Descuento;
                                }
                                double importeMenor = Importe / 0.85;
                                tabla.Rows.Add(Cantidad, Descripcion, NoIdentificacion, importeMenor);
                                subTotal += importeMenor;
                            }

                        }
                    }
                }
            }
            //Si el documento falla o tiene algun error de estructura cae al else
            catch (Exception)
            {
                MessageBox.Show("El documento XML no contine la extructura adecuada para ser procesado, esto es debido a que en su estructura, no se encuentra ninguna etiqueta padre con cfdi:Conceptos", "Error de estructura", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            //Operaciones del Subtotal, Iva y Total.
            tabla.Rows.Add("", "", "Subtotal", subTotal);
                iva = subTotal * 0.16;
                tabla.Rows.Add("", "Gracias por permitirnos servirle", "IVA", iva);
                totalGeneral = subTotal + iva;
                tabla.Rows.Add("", "", "Total", totalGeneral);
                tabla.Rows.Add("", "", "", "");
            //Pie del documento
                tabla.Rows.Add("", "Atentamente", "", "");
                tabla.Rows.Add("", "", "", "");
                tabla.Rows.Add("", "", "", "");
                tabla.Rows.Add("", "____________________________________________________", "", "");
                tabla.Rows.Add("", "Ing. Pedro Hernández González", "", "");
                tabla.Rows.Add("", "Administración de Servicios Carreteros", "", "");
                tabla.Rows.Add("", "", "", "");
                tabla.Rows.Add("", "", "", "");
                tabla.Rows.Add("", "____________________________________________________", "", "");
                tabla.Rows.Add("", "Firma de Autorización de Reparación y", "", "");
                tabla.Rows.Add("", "Nombre del Personal Autorizado", "", "");
                tabla.Rows.Add("", "", "", "");
                tabla.Rows.Add("", "", "", "");
                //tabla.Rows.Add("KENWORTH DEL ESTE S.A. DE C.V:B", "", "", "TEL. (01 271) 71 71400 EXT. 211");
                //tabla.Rows.Add("BANCOMER 0452279959", "", "", "C. INTER. BANCOMER 012855004522799591");
                //tabla.Rows.Add("BANAMEX 194864-3 SUC 815", "", "", "C. INTER BANAMEX 222855081519486439");
                style.SetWrapText(true); //Esto es para mandar hacia abajo el texto de la celda en caso de no caber.
            //Asignacion de estilos a las columnas.
                documento.SetColumnStyle(1, style);
                documento.SetColumnStyle(2, style);
                documento.SetColumnStyle(3, style);
                documento.SetColumnStyle(4, style);
                documento.ImportDataTable(7, 1, tabla, true);//Comenzamos la taba en la fila 7, columna 1, introducimos la tabla, y colocamos true
            //Intentamos guardar el documento
            try
            {
                SaveFileDialog dlg = new SaveFileDialog();
                dlg.Filter = "text file|*.xlsx"; 
                if (dlg.ShowDialog() == DialogResult.OK) 
                {
                    //Procedemos a guardar el documento en la direccion que el usuario coloque.
                    string pat = dlg.FileName;
                    documento.SaveAs(pat);
                }
        }
            catch (Exception)
            {

                MessageBox.Show("No se logro guardar el documento, debido a un error", "Error al guardar", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
            }
}
    }
}