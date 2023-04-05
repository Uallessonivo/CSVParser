using System.Data;
using System.Globalization;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using ExcelDataReader;

namespace CSVParser;

public class CsvRecords
{
    public long Serial { get; set; }
    public long Cpf { get; set; }
    public string? Valor { get; set; }
    public string? Colaborador { get; set; }
}

public sealed class CsvHeader : ClassMap<CsvRecords>
{
    public CsvHeader()
    {
        Map(c => c.Serial).Name("Numero de Serie");
        Map(c => c.Cpf).Name("CPF");
        Map(c => c.Valor).Name("Valor da Carga");
        Map(c => c.Colaborador).Name("Observacao");
    }
}

internal class Program
{
    private static void Main(string[] args)
    {
        var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Files", "cartoes.xlsx");

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

        var reader = ExcelReaderFactory.CreateReader(stream);

        var dataSet = reader.AsDataSet();

        var records = new List<CsvRecords>();

        foreach (var row in dataSet.Tables[0].Rows.Cast<DataRow>().Skip(1))
        {
            // Need to verify why is returning two columns with '0'

            records.Add(new CsvRecords
            {
                Serial = long.Parse((string)row[1]),
                Cpf = long.Parse(row[2].ToString()
                    .Replace(".", "")
                    .Replace("-", "")),
                Valor = "",
                Colaborador = row[3].ToString().Substring(0, Math.Min(row[3].ToString().Length, 35))
            });
        }

        Console.WriteLine("Digite o nome do arquivo a ser gerado: ");
        var filename = Console.ReadLine();

        var outputFilePath = Path.Combine(Directory.GetCurrentDirectory(), "Files",
            $"{DateTime.Today:yyyMMdd}-{filename}.csv");

        using var writer = new StreamWriter(outputFilePath);

        using var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = ";"
        });

        csv.Context.RegisterClassMap<CsvHeader>();
        csv.WriteRecords(new[] { new CsvRecords() });
        csv.WriteRecords(records);
    }
}