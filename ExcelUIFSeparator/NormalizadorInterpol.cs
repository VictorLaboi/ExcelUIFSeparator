using DocumentFormat.OpenXml.EMMA;
using ExcelUIFSeparator;
using System.Text;
using System.Text.RegularExpressions;

class NormalizadorInterpol
{
    public void Procesar(string rutaEntrada, string rutaSalidaSql)
    {
        var registros = File.ReadAllText(rutaEntrada)
                            .Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);

        var salida = new StringBuilder();

        foreach (var bloque in registros)
        {
            var matchId = Regex.Match(bloque, @"^(\d+)\|(.+)");
            if (!matchId.Success)
                continue;

            int id = int.Parse(matchId.Groups[1].Value);
            string nombrePrincipal = matchId.Groups[2].Value.Trim();
            nombrePrincipal = nombrePrincipal.Replace("'", "´");
            var texto = Regex.Replace(nombrePrincipal, @"\s+", " ").Trim();
            texto = texto.Normal();

            salida.AppendLine($"INSERT INTO EntidadesInfo(idEntidad, FechaNacimiento) values({id}, '{texto.Replace("°", null)}')");

            IEnumerable<string> nombresAlternos;
                nombresAlternos = bloque.Split('\n')
                           .Skip(1)
                           .Select(x => x.Trim().Replace("'", "´"))
                           .Where(x => !string.IsNullOrWhiteSpace(x) && !x.StartsWith(id + "|"));


            foreach (var alterno in nombresAlternos)
            {
                salida.AppendLine($"INSERT INTO EntidadesInfo(idEntidad, FechaNacimiento) values({id}, '{texto.Replace("°", null)}')");
            }
        }

        File.WriteAllText(rutaSalidaSql, salida.ToString());
    }

}
public static class StringExtensions
{
    public static string Normal(this string texto)
    {
        if (string.IsNullOrWhiteSpace(texto))
            return string.Empty;

        // Reemplaza múltiples espacios por uno solo
        string limpio = Regex.Replace(texto, @"\s+", " ");

        // Elimina espacios al inicio y al final
        return limpio.Trim();
    }
}


