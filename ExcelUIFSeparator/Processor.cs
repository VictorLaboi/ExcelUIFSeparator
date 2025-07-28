using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelUIFSeparator
{
    internal class Processor
    {
        public void Procesar(string rutaEntrada, string rutaSalidaCsv)
        {
            var registros = File.ReadAllText(rutaEntrada)
                                .Split('|', StringSplitOptions.RemoveEmptyEntries);

            var salida = new StringBuilder();
            salida.AppendLine("NombrePrincipal,NombresAlternos");

            foreach (var registro in registros)
            {
                string texto = registro.Replace("\r", " ").Replace("\n", " ").Trim();

                string nombrePrincipal = "";
                List<string> nombresAlternos = new();

                // Buscar la posición de "también conocido como", tolerando espacios
                var match = Regex.Match(texto, @"también\s+conocido\s+como", RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    int index = match.Index;

                    // Extraer el nombre principal antes de esa frase
                    nombrePrincipal = texto[..index].TrimEnd(',', ' ', '\t');

                    // Extraer la parte después de "también conocido como"
                    string alternosTexto = texto[(index + match.Length)..].Trim();

                    // Buscar con etiquetas (a), b), etc.)
                    var matches = Regex.Matches(alternosTexto, @"[a-zA-Z0-9]\)\s*([^\n\r,;]+)");
                    foreach (Match m in matches)
                    {
                        var nombreAlt = m.Groups[1].Value.Trim();
                        if (!string.IsNullOrWhiteSpace(nombreAlt))
                            nombresAlternos.Add(nombreAlt);
                    }

                    // Si no hubo etiquetas, dividir por comas o punto y coma
                    if (nombresAlternos.Count == 0)
                    {
                        var sinEtiquetas = alternosTexto.Replace("\r", " ").Replace("\n", " ");
                        nombresAlternos.AddRange(
                            sinEtiquetas.Split(';', ',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                        );
                    }
                }
                else
                {
                    // Si no hay "también conocido como", tomar el texto antes de la primera coma como nombre principal
                    var nombreSimple = Regex.Match(texto, @"^(.*?),");
                    if (nombreSimple.Success)
                    {
                        nombrePrincipal = nombreSimple.Groups[1].Value.Trim();
                    }
                    else
                    {
                        nombrePrincipal = texto;
                    }
                }

                salida.AppendLine($"\"{nombrePrincipal}\",\"{string.Join("▬ ", nombresAlternos)}\"");
            }

            using (StreamWriter writer = new StreamWriter(rutaSalidaCsv))
            {
                writer.WriteLine(salida.ToString());
            }
        }
    }
}
