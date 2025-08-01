using System.Text;
using System.Text.RegularExpressions;

class NormalizadorInterpol
{
    public void Procesar(string rutaEntrada, string rutaSalidaCsv)
    {
        var registros = File.ReadAllText(rutaEntrada)
                            .Split(';', StringSplitOptions.RemoveEmptyEntries);

        var salida = new StringBuilder();
        salida.AppendLine("NombrePrincipal,NombresAlternos");

        foreach (var registro in registros)
        {
            var texto = registro.Replace("\r", " ").Replace("\n", " ").Trim();
            string numeroCurrent = texto.Split("|")[0];

            string nombrePrincipal = ExtraerNombrePrincipal(texto[(texto.IndexOf("|") + 1)..]);

            List<string> nombresAlternos = new();

            

            string pattern = @"([a-z0-9])\)\s+(.+?)(?=(?:[a-z0-9]\)\s+)|$)";

            var matches = Regex.Matches(texto, pattern, RegexOptions.IgnoreCase);

            foreach (Match match in matches)
            {
                string nombre = match.Groups[2].Value.Trim();
                if (!string.IsNullOrWhiteSpace(nombre))
                    nombresAlternos.Add(nombre);
            }


            salida.AppendLine(
                $"--{numeroCurrent}:{nombrePrincipal}\n" +
                string.Join("\n", nombresAlternos.Select(n =>
                    $"INSERT INTO alt_entidad (idEntidad, NombreAlternativo) SELECT Id, '{n.Replace("'", "`").Trim()}' FROM UIFEntidades WHERE Nombre COLLATE Latin1_General_CI_AI LIKE '%{nombrePrincipal.Replace("'", "''").Trim()}%';"
                ))
            );
        }

        File.WriteAllText(rutaSalidaCsv, salida.ToString());
    }

    private string ExtraerNombrePrincipal(string texto)
    {
        // Si hay número y '|', lo eliminamos
        var separador = texto.IndexOf('|');
        if (separador >= 0 && separador + 1 < texto.Length)
            texto = texto[(separador + 1)..].Trim();

        // Buscar índice de cualquier patrón de "también conocido como"
        var patrones = new[]
        {
        "también conocido como",
        ", también conocido como",
        ",también conocido como"
    };

        int index = -1;
        foreach (var patron in patrones)
        {
            index = texto.IndexOf(patron, StringComparison.OrdinalIgnoreCase);
            if (index >= 0)
                break;
        }

        // Extraer parte principal
        string principal = index >= 0 ? texto[..index].Trim() : texto.Trim();

        // Quitar coma final si existe
        if (principal.EndsWith(","))
            principal = principal[..^1].Trim();

        return SanearNombre(principal);
    }


    private string SanearNombre(string nombre)
    {
        return Regex.Replace(nombre, @"\s+", " "); // quita espacios dobles
    }
}
