using DocumentFormat.OpenXml.Wordprocessing;
using ExcelDataReader;
using System.Data;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using UglyToad.PdfPig;
namespace PLDMexicoMejorado.Utils
{
    public class ExcelHandler
    {
        public List<string> GenerarListaDesdeExcel<T>(Stream streamXlsx, string? propiedad, params string[] propParams)
        {

            /*
             Donde matches omitidos son todos aquellos parámetros que el usuario busca omitir.
            Se transforman a Regex y despues se pasan como parámetro por medio de un foreach a los criterios de omision 
             */



            List<string> resultado = new List<string>();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            try
            {
                using (var reader = ExcelReaderFactory.CreateReader(streamXlsx))
                {
                    var dataset = reader.AsDataSet(new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = true, // Usa la primera fila como nombres de columna
                            FilterColumn = (columnReader, columnIndex) => columnIndex == 0 && propParams[0] != null || propParams[0] != string.Empty 
                            || columnIndex == 1 && propParams[1] != null || propParams[1] != string.Empty
                            //ReadHeaderRow = rowReader => { /* lógica para leer encabezados si necesitas */ }
                        }
                    });
                    
                    DataTable table = dataset.Tables[0];

                    foreach (DataRow row in table.Rows)
                    {
                        T instancia = (T)Activator.CreateInstance(typeof(T));
                        foreach (var prop in typeof(T).GetProperties())
                        {
                            if (table.Columns.Contains(prop.Name))
                            {
                                object valor = row[prop.Name];
                                if (valor != null && valor != DBNull.Value)
                                {
                                    prop.SetValue(instancia, Convert.ChangeType(valor, prop.PropertyType));
                                }
                            }
                        }


                        /*
                         A partir de este punto es donde se comenzará a hacer la relevacion entre params establecidos, bajo la condicion de que
                        dentro de la tabla Excel realmente se encuentren esas columanas, de otro modo lo establecido será string.Empty o null. 

                         */
                        foreach (var valueProp in propParams) //Obtenemos los parametros de la clase T
                        {
                            var propiedadInfo = instancia.GetType().GetProperty(valueProp); //Aqui es donde irá el var del foreach.
                            if (propiedadInfo != null)
                            {
                                /*La instancia establece 2 valores que hacen match con lo establecido dentro de propValues.
                                Cuando no coinciden, no se procede. 
                                 */

                                var val = propiedadInfo.GetValue(instancia);


                                /*Faltan las condicionantes donde "FechaNacimiento" o cualquier otro establecido dentro del DTO se cumple
                                 
                                 La iteracion solo valida el primer parametro del DTO pero no hace validaciones si dentro del DTO.Params > 1 || <0
                                Por lo que es necesario almacenar este valor dentro de un array para poder iterarlo como es debido o en su defecto, hacer que vaya 1:1 SIN repetir.
                                 
                                 */




                                if (val != null)
                                {
                                    string texto = val.ToString();

                                    if (propiedadInfo.Name == "F")
                                    {
                                        int indexer = texto.IndexOf(")");
                                        if (indexer > 0 && indexer < texto.Length - 1)
                                        {
                                            texto = texto.Substring(indexer + 1).Trim();
                                        }
                                        resultado.Add(texto);
                                    }

                                    //resultado.Add(texto);
                                    if (propiedadInfo.Name == "Principal")
                                    {
                                        // Obtener solo lo anterior a "también conocido como" (insensible a mayúsculas)
                                        var regex = new Regex(@"t[aá]mb[ií]?[eé]n\s+conoc[ií]do\s+como[:：]", RegexOptions.IgnoreCase);
                                        var match = regex.Match(texto);
                                        if (match.Success)
                                        {
                                            texto = texto.Substring(0, match.Index).Trim();
                                        }
                                        int interpol = texto.IndexOf("Liga de", StringComparison.OrdinalIgnoreCase);
                                        if (interpol > 0)
                                        {
                                            texto = texto.Substring(0, interpol).Trim();
                                        }
                                        // Solo agregar si NO contiene lo no deseado
                                        if (!texto.Contains("Página", StringComparison.OrdinalIgnoreCase)
                                            && !texto.Contains("www.gob.mx/uif", StringComparison.OrdinalIgnoreCase)
                                            && !texto.Contains("https", StringComparison.OrdinalIgnoreCase))
                                        {
                                            resultado.Add(texto);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                return resultado;
            }
            catch (Exception ex)
            {
                // Puedes loggear el error si quieres trazabilidad
                return new List<string>();
            }
        }

        public List<string> GenerarXML(Stream streamFile, string propiedad)
        {
            List<string> nodeValue = new List<string>();
            XmlDocument doc = new XmlDocument();
            if (streamFile != null)
            {
                try
                {
                    doc.Load(streamFile);
                }
                catch (Exception e)
                { }

                XmlNodeList nodeList = doc.GetElementsByTagName(propiedad);
                foreach (XmlNode node in nodeList)
                {
                    string value = node.InnerText.Trim();
                    nodeValue.Add(value);
                }

                return nodeValue;
            }
            return nodeValue;
        }
   
    public List<string> GenerarListaDesdePDFDesdeNombre(Stream pdfStream, string marcadorInicio = "Nombre", Func<string, bool> filtro = null)
        {
            var resultado = new List<string>();
            bool iniciarCaptura = false;

            using (var reader = PdfDocument.Open(pdfStream))
            {
                foreach (var pagina in reader.GetPages())
                {
                    string[] lineas = pagina.Text.Split('\n');
                    foreach (var linea in lineas)
                    {
                        string valor = linea.Trim();

                        if (string.IsNullOrWhiteSpace(valor))
                            continue;

                        if (!iniciarCaptura && valor.Equals(marcadorInicio, StringComparison.OrdinalIgnoreCase))
                        {
                            iniciarCaptura = true;
                            continue;
                        }

                        if (iniciarCaptura)
                        {
                            if (filtro == null || filtro(valor))
                                resultado.Add(valor);
                        }
                    }
                }
            }

            return resultado;
        }
    }
}