namespace Cognitive.ExcelUtility
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Please wait while the report is generated...");
            var workbook = ExcelBuilder.CreateUserReadableReport().Result;
            var path = $"{Directory.GetCurrentDirectory()}\\reports";
            Directory.CreateDirectory(path);
            var filename = $@"report_{DateTime.Now.ToString("yy-MM-dd_HH-mm-ss")}.xlsx";
            using (FileStream fileStream = new FileStream($"{path}\\{filename}", FileMode.Create))
            {
                workbook.Write(fileStream);
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"The report is generated in {path}\\{filename})");
            Console.ReadKey();
        }
    }
}
