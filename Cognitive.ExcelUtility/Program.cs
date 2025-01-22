using NPOI.SS.UserModel;

namespace Cognitive.ExcelUtility
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            var workbook = ExcelBuilder.CreateDetalisedReport().Result;
            using (FileStream fileStream = new FileStream("C:\\Users\\x1larus\\Desktop\\file.xlsx", FileMode.Create))
            {
                workbook.Write(fileStream);
            }
        }
    }
}
