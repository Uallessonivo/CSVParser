using ExcelDataReader;

namespace CSVParser;

internal class Program
{
    private static void Main(string[] args)
    {
        string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Files", "cartoes.xlsx");

        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                do
                {
                    while (reader.Read())
                    {
                        Console.WriteLine(reader.GetValue(1));
                        Console.WriteLine(reader.GetValue(2));
                        Console.WriteLine(reader.GetValue(3));
                        Console.WriteLine(reader.GetValue(4));
                        Console.WriteLine(reader.GetValue(5));
                        Console.WriteLine(reader.GetValue(6));
                    }
                } while (reader.NextResult());
            }
        }
    }
}