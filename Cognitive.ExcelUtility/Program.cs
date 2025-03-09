using System.Text;
using Cognitive.ExcelUtility.DalPg.Base;

namespace Cognitive.ExcelUtility
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            var isDev = false;// args.Contains("--dev");
            var isDataset = true; //args.Contains("--dataset");
            
            DbHelper.BuildDataSource(isDev);

            Console.WriteLine("Please wait while the report is generated...");
            
            var path = BuildPath(isDev, isDataset);

            if (isDataset)
            {
                var csv = CsvBuilder.BuildDataset();
                var bytes = Encoding.UTF8.GetBytes(csv);
                using var fileStream = new FileStream(path, FileMode.Create);
                fileStream.Write(bytes, 0, bytes.Length);
            }
            else
            {
                var workbook = ExcelBuilder.CreateUserReadableReport().Result;
                using var fileStream = new FileStream(path, FileMode.Create);
                workbook.Write(fileStream);
            }

            Console.WriteLine($"The report is generated in");
            Console.WriteLine($"{path}");
            Console.WriteLine($"Press any key to continue...");
            Console.Read();
        }

        private static string BuildPath(bool isDev, bool isDataset)
        {
            var path = $"{Directory.GetCurrentDirectory()}{Path.DirectorySeparatorChar}reports";
            Directory.CreateDirectory(path);

            var sb = new StringBuilder($"{path}{Path.DirectorySeparatorChar}");
            sb.Append(isDev ? "DEV-" : "REL-");
            sb.Append(isDataset ? "dataset_" : "report_");
            sb.Append(DateTime.Now.ToString("yy-MM-dd_HH-mm-ss"));
            sb.Append(isDataset ? ".csv" : ".xlsx");
            return sb.ToString();
        }
    }
}
