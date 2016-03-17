using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Printing;
using System.Xml;
using System.IO;
using System.Security.Permissions;
using System.Globalization;

namespace IMP
{
    class Program
    {
        private static Boolean depuracion = false;

        private static XmlDocument xConfig = new XmlDocument();
        private static XmlDocument xTrabajo = new XmlDocument();
        private static XmlDocument xPlantilla = new XmlDocument();

        private static string cfg_r_plantillas = "";
        private static string cfg_r_supervision = "";
        private static string cfg_r_respaldo = "";

        public static System.Drawing.Printing.PrintDocument prnDocument;
        private static Font Fuente = new Font("Arial", 8, FontStyle.Regular);
        private static SolidBrush Brush = new SolidBrush(Color.Black);

        public static float PAL = 0F;
        public static float PAT = 0F;

        public static float rellenoIzq = 0F;
        public static float rellenoSup = 0F;

        public static string proy_titulo = "7G X-Labs: IMP 0.1 - Impresión por Medio de Plantillas";
        private static FileSystemWatcher watcher = new FileSystemWatcher();

        static void Main(string[] args)
        {
            Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture;
            AppDomain.CurrentDomain.UnhandledException += UnhandledExceptionTrapper;

            CargarConfiguracion();
            Console.Title = proy_titulo;
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.BackgroundColor = ConsoleColor.Black;
            
            Console.WriteLine(proy_titulo);
            Console.WriteLine("* Ruta.Plantillas = " + cfg_r_plantillas);
            Console.WriteLine("* Ruta.Supervisión = " + cfg_r_supervision);
            Console.WriteLine("* Ruta.Respaldo = " + cfg_r_respaldo);

            InicializarComponentesDeImpresion();
            InicializarSupervisorDeTrabajos();

            Console.WriteLine("* Aceptando trabajos\n");
            Console.ReadLine();
        }

        static void UnhandledExceptionTrapper(object sender, UnhandledExceptionEventArgs e)
        {
            Console.WriteLine(e.ExceptionObject.ToString());
            Console.WriteLine("Ocurrió una super excepción. Presione Enter para terminar.");
            Console.ReadLine();
            Environment.Exit(1);
        }

        private static void prnDocument_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.PageUnit = GraphicsUnit.Inch;
            PAL = (e.PageSettings.PrintableArea.Left / 100);
            PAT = (e.PageSettings.PrintableArea.Top / 100);

            if (depuracion == true)
            {
                e.Graphics.DrawString("Margin Izq:" + e.MarginBounds.Left, Fuente, Brush, (float)1, (float)2);
                e.Graphics.DrawString("Margin Sup:" + e.MarginBounds.Top, Fuente, Brush, (float)1, (float)2.5);
                e.Graphics.DrawString("Lim Izq:" + e.PageSettings.PrintableArea.Left, Fuente, Brush, (float)1, (float)3);
                e.Graphics.DrawString("Lim Sup:" + e.PageSettings.PrintableArea.Top, Fuente, Brush, (float)1, (float)3.5);
                e.Graphics.DrawString("Pag Alto:" + e.MarginBounds.Height, Fuente, Brush, (float)1, (float)4);
                e.Graphics.DrawString("Pag Ancho:" + e.MarginBounds.Width, Fuente, Brush, (float)1, (float)4.5);
            }

            DocumentoProcesarCampos(e);
            DocumentoProcesarTablas(e);

        }

        private static void DocumentoProcesarTablas(PrintPageEventArgs e)
        {
            float x = 0F;
            float y = 0F;
            float interlineado = 0F;

            XmlNodeList xml_tablas = xTrabajo["impresion"]["datos"].GetElementsByTagName("tabla");
            for (int i = 0; i < xml_tablas.Count; i++)
            {
                string id_tabla = xml_tablas[i].Attributes["id"].Value;
                y = float.Parse(xPlantilla.SelectSingleNode("/plantilla/contenido/tabla[@id='" + id_tabla + "']").Attributes["y"].Value);
                interlineado = float.Parse(xPlantilla.SelectSingleNode("/plantilla/contenido/tabla[@id='" + id_tabla + "']").Attributes["interlineado"].Value);

                XmlNodeList xml_filas = xml_tablas[i].ChildNodes;
                for (int f = 0; f < xml_filas.Count; f++)
                {
                    XmlNodeList xml_columnas = xml_filas[f].ChildNodes;
                    for (int c = 0; c < xml_columnas.Count; c++)
                    {
                        x = float.Parse(xPlantilla.SelectSingleNode("/plantilla/contenido/tabla[@id='" + id_tabla + "']/columna[" + (c+1) + "]").Attributes["x"].Value);

                        ImprimirCampoEnCanvas(e, xml_columnas[c].InnerText, x, y);
                    }

                    y += interlineado;
                }
            }
            
        }

        private static void DocumentoProcesarCampos(PrintPageEventArgs e)
        {
            // Hay que iterar el documento por los campos e ir buscando en la plantilla sus atributos
            XmlNodeList elemList = xTrabajo["impresion"]["datos"].GetElementsByTagName("campo");
            for (int i = 0; i < elemList.Count; i++)
            {

                float x = 0F;
                float y = 0F;
                string id = "";
                string texto = "";

                texto = elemList[i].InnerText;

                if (texto == "") continue;

                id = elemList[i].Attributes["id"].Value;
                
                if (xPlantilla.SelectNodes("/plantilla/contenido/campo[@id='" + id + "']").Count == 0) continue;

                x = float.Parse(xPlantilla.SelectSingleNode("/plantilla/contenido/campo[@id='" + id + "']").Attributes["x"].Value);
                y = float.Parse(xPlantilla.SelectSingleNode("/plantilla/contenido/campo[@id='" + id + "']").Attributes["y"].Value);

                

                if (depuracion == true)
                {
                    Console.WriteLine("[" + id + "]" + texto + " :: " + xPlantilla.SelectSingleNode("/plantilla/contenido/campo[@id='" + id + "']").Attributes["x"].Value + "x" + xPlantilla.SelectSingleNode("/plantilla/contenido/campo[@id='" + id + "']").Attributes["y"].Value);
                    Console.WriteLine("[" + id + "]" + texto + " :: " + x + "x" + y);
                }

                ImprimirCampoEnCanvas(e, texto, x, y);
                
            }
        }

        private static void ImprimirCampoEnCanvas(PrintPageEventArgs e, string texto, float x, float y)
        {
            float a = e.Graphics.MeasureString(texto, Fuente).Height;
            e.Graphics.DrawString(texto, Fuente, Brush, ((x - PAL) + rellenoIzq), (((y - a) - PAT)) + rellenoSup);
        }

        static void InicializarComponentesDeImpresion()
        {
            prnDocument = new System.Drawing.Printing.PrintDocument();
            prnDocument.PrintController = new StandardPrintController();
            prnDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(prnDocument_PrintPage);
        }

        static void InicializarSupervisorDeTrabajos()
        {
            watcher.Path = cfg_r_supervision;
            watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;
            watcher.Filter = "*.xml";
            watcher.Created += new FileSystemEventHandler(CrearNuevoTrabajo);
            watcher.EnableRaisingEvents = true;
        }

        private static void CrearNuevoTrabajo(object sender, FileSystemEventArgs e)
        {
            string plantilla = "";
            string revision = "";
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("Trabajo iniciado: " + e.FullPath);
            Console.ForegroundColor = ConsoleColor.DarkMagenta;

            xTrabajo.Load(e.FullPath);
            plantilla = xTrabajo["impresion"]["documento"]["plantilla"].InnerText;
            revision = xTrabajo["impresion"]["documento"]["revision"].InnerText;
            Console.WriteLine("Plantilla: " + plantilla + " - Rev. " + revision);
            plantilla = cfg_r_plantillas + plantilla + ".xml";
            if (!File.Exists(plantilla))
            {
                Console.WriteLine("Plantilla NO existe - cancelando");
            } else {                
                xPlantilla.Load(plantilla);
                Console.WriteLine("Tenemos " + xPlantilla["plantilla"]["documento"]["descripcion"].InnerText + " - Rev. " + xPlantilla["plantilla"]["documento"]["revision"].InnerText);

                prnDocument.PrinterSettings.PrinterName = xPlantilla["plantilla"]["documento"]["impresora"].InnerText;
                prnDocument.DefaultPageSettings.PaperSize = new PaperSize("Custom", int.Parse(xPlantilla["plantilla"]["documento"]["impresora"].Attributes["ancho"].Value), int.Parse(xPlantilla["plantilla"]["documento"]["impresora"].Attributes["alto"].Value));
                prnDocument.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);

                string fuenteDeseada = xPlantilla["plantilla"]["documento"]["impresora"].Attributes["fuente"].Value;
                float fuenteTamanoDeseado = float.Parse(xPlantilla["plantilla"]["documento"]["impresora"].Attributes["fuenteTamano"].Value);

                if (fuenteDeseada != "" && fuenteTamanoDeseado != 0)
                {
                    Fuente = new Font(fuenteDeseada, fuenteTamanoDeseado, FontStyle.Regular);
                }

                rellenoIzq = float.Parse(xPlantilla["plantilla"]["documento"]["impresora"].Attributes["rellenoIzq"].Value);
                rellenoSup = float.Parse(xPlantilla["plantilla"]["documento"]["impresora"].Attributes["rellenoSup"].Value);
                prnDocument.Print();
            }
            
            File.Delete(e.FullPath);
            
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("Trabajo completado: " + e.Name);
            Console.ForegroundColor = ConsoleColor.Yellow;
        }

        static void CargarConfiguracion()
        {
            xConfig.Load("./config.xml");
            cfg_r_plantillas = xConfig.GetElementById("plantillas").InnerText;
            cfg_r_supervision = xConfig.GetElementById("supervision").InnerText;
            cfg_r_respaldo = xConfig.GetElementById("respaldos").InnerText;
            depuracion = ( xConfig["configuracion"]["depuracion"].InnerText == "si" ? true : false ) ;

        }
    }
}
