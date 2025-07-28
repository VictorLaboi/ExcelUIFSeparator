using ExcelUIFSeparator;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using PLDMexicoMejorado.Utils;
using System.Xml.XPath;

ExcelHandler handler = new ExcelHandler();

string path = @"C:\Users\Lione\OneDrive\Escritorio\LISTAS ACTUALIZADAS 2025";
List<string> names = GetDocumentNames(@"C:\Users\Lione\OneDrive\Escritorio\LISTAS ACTUALIZADAS 2025");

string[] criteriosOmitidos = new[]
{
    "también conocido como", // patrón con acentos y variantes: @"t[aá]mb[ií]?[eé]n\s+conoc[ií]do\s+como[:：]"
    "Liga de",               // usado como punto de corte (IndexOf)
    "Página",                // filtrado por contenido no deseado
    "www.gob.mx/uif",        // URL gubernamental filtrada
    "https"                  // contenido web filtrado
};
var streams = GetExcelStreams(path, names);

List<List<string>> prop = new List<List<string>>();
foreach (var stream in streams) 
{
    prop.Add(handler.GenerarListaDesdeExcel<DataSecondColumn>(stream, "","Principal","F"));
}

if (!prop.Any()) 
{
    Console.WriteLine("No se han encontrado elementos");
    return;
}
int contador = 1;
string outputFolder = @"C:\Users\Lione\Downloads\UIF";

foreach (var writeCurrent in prop) //This is the best code i´ve ever writen
{
    string outputPath = Path.Combine(outputFolder, $"Documento_{contador}.txt");

    using (StreamWriter writer = new StreamWriter(outputPath))
    {
        foreach (var linea in writeCurrent)
        {
            writer.WriteLine(linea+"|");
        }
    }

    contador++;
}

NormalizadorInterpol normalizador = new NormalizadorInterpol();
List<string> Nombres = GetDocumentNames($"{outputFolder}");
List<string> SalidaDatos = new List<string>();

foreach (var file in Nombres)
{
    string nuevoStr = Path.Combine(outputFolder, file);
    if (nuevoStr is not null) 
    {
        normalizador.Procesar($"{nuevoStr}", $@"C:\Users\Lione\Downloads\UIF\Normalizado\{file}");
    }
}


static List<FileStream> GetExcelStreams(string folderPath, List<string> fileNames)
{
    string[] ExcelExtensions = { ".xls", ".xlsx" };
    var excelStreams = new List<FileStream>();

    foreach (var fileName in fileNames)
    {
        var extension = Path.GetExtension(fileName);
        if (Array.Exists(ExcelExtensions, ext => ext.Equals(extension, StringComparison.OrdinalIgnoreCase)))
        {
            string fullPath = Path.Combine(folderPath, fileName);
            try
            {
                var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                excelStreams.Add(stream);
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Error al abrir '{fileName}': {ex.Message}");
            }
        }
    }

    return excelStreams;
}

static List<string> GetDocumentNames(string folderPath)
{
 string[] DocumentExtensions = { ".doc", ".docx", ".pdf", ".txt", ".xls", ".xlsx" };

if (string.IsNullOrWhiteSpace(folderPath))
        throw new ArgumentException("La ruta no puede estar vacía.", nameof(folderPath));

    if (!Directory.Exists(folderPath))
        throw new DirectoryNotFoundException($"La carpeta '{folderPath}' no existe.");

    var documentNames = new List<string>();

    foreach (var file in Directory.EnumerateFiles(folderPath))
    {
        var extension = Path.GetExtension(file);
        if (Array.Exists(DocumentExtensions, ext => ext.Equals(extension, StringComparison.OrdinalIgnoreCase)))
        {
            documentNames.Add(Path.GetFileName(file));
        }
    }

    return documentNames;
}
