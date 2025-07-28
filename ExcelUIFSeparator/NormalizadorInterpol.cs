using System.Text.RegularExpressions;
using System.Text;
using System.IO;

class NormalizadorInterpol
{
    public void Procesar(string rutaEntrada, string rutaSalidaCsv)
    {
        var registros = File.ReadAllText(rutaEntrada)
                            .Split('|', StringSplitOptions.RemoveEmptyEntries);

        var salida = new StringBuilder();
        salida.AppendLine("NombrePrincipal,NombresAlternos");

        foreach (var registro in registros)
        {
            var texto = registro.Replace("\r", " ").Replace("\n", " ").Trim();

            string nombrePrincipal = "";
            List<string> nombresAlternos = new();

            // Extraer nombre principal (hasta antes de "también conocido como")
            var nombrePrincipalMatch = Regex.Match(texto, @"^(.*?)(?=\s*,?\s*también conocido como)", RegexOptions.IgnoreCase);
            if (nombrePrincipalMatch.Success)
            {
                nombrePrincipal = nombrePrincipalMatch.Groups[1].Value.Trim().TrimEnd(',');
            }
            else
            {
                // Si no hay "también conocido como", tomar el texto hasta la coma final
                var simpleNombre = Regex.Match(texto, @"^(.*?),");
                if (simpleNombre.Success)
                {
                    nombrePrincipal = simpleNombre.Groups[1].Value.Trim();
                }
                else
                {
                    nombrePrincipal = texto; // último recurso
                }
            }

            // Extraer nombres alternos
            var alternosMatch = Regex.Match(texto, @"también conocido como[:\s]*(.*)", RegexOptions.IgnoreCase);
            if (alternosMatch.Success)
            {
                var alternosTexto = alternosMatch.Groups[1].Value;

                // Buscar si vienen con a), b), etc.
                var matches = Regex.Matches(alternosTexto, @"[a-zA-Z0-9]\)\s*([^\n\r,;]+)");
                foreach (Match m in matches)
                {
                    var nombreAlt = m.Groups[1].Value.Trim();
                    if (!string.IsNullOrWhiteSpace(nombreAlt))
                        nombresAlternos.Add(nombreAlt);
                }

                // Si no hay con etiquetas, dividir por saltos o comas
                if (nombresAlternos.Count == 0)
                {
                    var sinEtiquetas = alternosTexto.Replace("\n", " ").Replace("\r", " ");
                    nombresAlternos.AddRange(
                        sinEtiquetas.Split(';', ',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    );
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
