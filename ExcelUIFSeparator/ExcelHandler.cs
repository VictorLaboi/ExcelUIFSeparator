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
        public List<string> GenerarListaDesdeExcel<T>(
          Stream streamXlsx,
          string? propiedad,
          out int contadorFinal,
          int contadorInicial,
         params string[] propParams)
        {

            int ContadorCm = contadorInicial;
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
                            UseHeaderRow = true,
                            FilterColumn = (columnReader, columnIndex) =>
                                columnIndex == 0 && !string.IsNullOrEmpty(propParams[0])
                                || columnIndex == 1 && !string.IsNullOrEmpty(propParams[1])
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

                        foreach (var valueProp in propParams)
                        {
                            var propiedadInfo = instancia.GetType().GetProperty(valueProp);
                            if (propiedadInfo != null)
                            {
                                var val = propiedadInfo.GetValue(instancia);
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

                                        resultado.Add($"{ContadorCm}|{texto}");
                                    }

                                    if (propiedadInfo.Name == "Principal")
                                    {
                                        var OnitRegex = new Regex(@"\b\d+[\.\-]?\s*['‘“\""]*\s*[A-Za-z]", RegexOptions.IgnoreCase);
                                        var matchOnInit = OnitRegex.Match(texto ?? string.Empty);
                                        if (matchOnInit.Success)
                                        {
                                            texto = Regex.Replace(texto ?? string.Empty, @"^\s*\d+[\.\-]?\s*", "").Trim();

                                            Regex[] regexRegulares = new Regex[] 
                                            {
                                                new Regex(@"[,\.]?\s*t[aá]mb[ií]?[eé]n\s+conoc[ií]do\s+como[:：]?", RegexOptions.IgnoreCase),
                                                new Regex(@"[,\.]?\s*t[aá]mb[ií][eé]n\b", RegexOptions.IgnoreCase),
                                                new Regex(@"\b[cC][oó][mM][oÓ]\b", RegexOptions.IgnoreCase),
                                                new Regex(@"\b[cC][oó][nN][oó][cC][ií][dD][oÓ]\b", RegexOptions.IgnoreCase),
                                                new Regex(@"(?<=\bconocida)([a-zA-Z]\))", RegexOptions.IgnoreCase),
                                            };

                                            foreach (Regex regulares in regexRegulares)
                                            {
                                                if(regulares.IsMatch(texto))
                                                {
                                                    texto = regulares.Replace(texto, string.Empty);
                                                }
                                            }
                                            Regex regexLetras = new Regex(@"(?<=[a-zA-Z\s,\.])([a-zA-Z]\))", RegexOptions.IgnoreCase);
                                            texto = regexLetras.Replace(texto, "\n");


                                            int interpol = texto.IndexOf("Liga de", StringComparison.OrdinalIgnoreCase);
                                            if (interpol > 0)
                                            {
                                                texto = texto.Substring(0, interpol).Trim();
                                            }

                                            if (!string.IsNullOrWhiteSpace(texto)
                                                && !texto.Contains("Página", StringComparison.OrdinalIgnoreCase)
                                                && !texto.Contains("www.gob.mx/uif", StringComparison.OrdinalIgnoreCase)
                                                && !texto.Contains("https", StringComparison.OrdinalIgnoreCase))
                                            {
                                                ContadorCm++;
                                                //resultado.Add($"{ContadorCm}|{texto}");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    contadorFinal = ContadorCm;
                    return resultado;
                }
            }
            catch
            {
                contadorFinal = ContadorCm;
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