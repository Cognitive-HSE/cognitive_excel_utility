using Cognitive.ExcelUtility.DalPg.Entities;
using Cognitive.ExcelUtility.DalPg.Managers;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Runtime.CompilerServices;

namespace Cognitive.ExcelUtility
{
    internal static class ExcelBuilder
    {
        public static async Task<IWorkbook> CreateDetalisedReport()
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("report");

            var data = await GetDetalisedData();

            var rownum = 0;
            foreach (var item in data)
            {
                var row = sheet.CreateRow(rownum++);

                row.CreateAndWriteCell(0, item.UserId);
                row.CreateAndWriteCell(1, item.TestId);
                row.CreateAndWriteCell(2, item.TryNumber);
                row.CreateAndWriteCell(3, item.Accuracy);
                row.CreateAndWriteCell(4, item.CompleteTime);
                row.CreateAndWriteCell(5, item.NumberCorrectAnswers);
                row.CreateAndWriteCell(6, item.NumberAllAnswers);
            }

            return workbook;
        }

        public static async Task<List<TestResultEntity>> GetDetalisedData()
        {
            return await new TestResultManager().ReadAll();
        }

        public static void CreateAndWriteCell(this IRow row, int colnum, object? val)
        {
            if (val == null) return;

            var cell = row.CreateCell(colnum);

            if (val is long || val is int || val is decimal)
            {
                cell.SetCellValue(Convert.ToDouble(val));
                return;
            }

            if (val is TimeSpan ts)
            {
                cell.SetCellValue(ts.TotalSeconds);
            }
        }
    }
}
 