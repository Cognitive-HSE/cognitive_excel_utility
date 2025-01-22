using Cognitive.ExcelUtility.DalPg.Entities;
using Cognitive.ExcelUtility.DalPg.Managers;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Cognitive.ExcelUtility
{
    internal static class ExcelBuilder
    {
        private static XSSFWorkbook _workbook = new XSSFWorkbook();
        private static ISheet _sheet = _workbook.CreateSheet("Report");

        #region User Readable Report

        public static async Task<IWorkbook> CreateUserReadableReport()
        {
            var headerNames = await GetHeaderNames();
            IRow[] headerRows = { _sheet.CreateRow(0), _sheet.CreateRow(1) };
            _sheet.CreateFreezePane(1, 2);

            // Сборка шапки юзеров
            headerRows[0].CreateAndWriteCell(0, "ИД пользователя", CellStyleHeader, 15);
            _sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 0));

            // Сборка шапки тестов
            for (int i = 0; i < headerNames.Count; i++)
            {
                var startRegionCol = 1 + i * 5; // 1 == количество столбцов с инфой о юзерах
                headerNames[i].StartRegionCol = startRegionCol;
                headerRows[0].CreateAndWriteCell(startRegionCol, headerNames[i].TestName, CellStyleHeader);
                _sheet.AddMergedRegion(new CellRangeAddress(0, 0, startRegionCol, startRegionCol + 4));

                headerRows[1].CreateAndWriteCell(startRegionCol + 0, "#", CellStyleHeader, 3);
                headerRows[1].CreateAndWriteCell(startRegionCol + 1, "% выполнения", CellStyleHeader, 12);
                headerRows[1].CreateAndWriteCell(startRegionCol + 2, "Время выполнения", CellStyleHeader, 12);
                headerRows[1].CreateAndWriteCell(startRegionCol + 3, headerNames[i].TitleCorrect, CellStyleHeader, 20);
                headerRows[1].CreateAndWriteCell(startRegionCol + 4, headerNames[i].TitleAll, CellStyleHeader, 20);
            }

            // Наполнение таблицы
            var data = await GetTestResults();
            var rownum = 2; // 2 == высота шапки

            foreach (var user in data.GroupBy(el => el.UserId))
            {
                var rowsCreated = 0;
                foreach (var testTry in user.GroupBy(el => el.TryNumber))
                {
                    var row = _sheet.CreateRow(rownum++);
                    rowsCreated++;
                    if (rowsCreated == 1)
                    {
                        row.CreateAndWriteCell(0, user.Key, CellStyleHeader);
                    }

                    foreach (var testRes in testTry)
                        WriteTestResult(row, headerNames.FirstOrDefault(el => el.TestId == testRes.TestId).StartRegionCol, testRes);
                }
                
                if (rowsCreated > 1)
                {
                    _sheet.AddMergedRegion(new CellRangeAddress(rownum - rowsCreated, rownum - 1, 0, 0));
                }
            }

            return _workbook;
        }

        private static void WriteTestResult(IRow row, int startCol, TestResultEntity item)
        {
            row.CreateAndWriteCell(startCol + 0, item.TryNumber);
            row.CreateAndWriteCell(startCol + 1, item.Accuracy);
            row.CreateAndWriteCell(startCol + 2, item.CompleteTime);
            row.CreateAndWriteCell(startCol + 3, item.NumberCorrectAnswers);
            row.CreateAndWriteCell(startCol + 4, item.NumberAllAnswers);
        }

        #endregion

        public static async Task<IWorkbook> CreateReportForTraining()
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("report");

            var data = await GetTestResults();

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

        #region Styles

        private static ICellStyle CellStyleHeader = CellStyleHeaderBuild();

        private static ICellStyle CellStyleHeaderBuild()
        {
            var style = _workbook.CreateCellStyle();

            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.WrapText = true;

            return style;
        }

        #endregion

        #region GetData

        private static async Task<List<TestResultEntity>> GetTestResults()
        {
            return await new TestResultManager().ReadAll();
        }

        private static async Task<List<TestNsiEntity>> GetHeaderNames()
        {
            return await new TestNsiManager().ReadAll();
        }

        #endregion

        #region Private

        /// <summary>
        /// Создает и пишет ячейку
        /// </summary>
        /// <param name="row">Строка для записи</param>
        /// <param name="colnum">Номер колонки</param>
        /// <param name="val">Значение</param>
        /// <param name="style">Стиль ячейки</param>
        /// <param name="colWidth">Ширина столбца</param>
        private static void CreateAndWriteCell(this IRow row, int colnum, object? val, ICellStyle? style = null, int? colWidth = null)
        {
            if (val == null) return;

            var cell = row.CreateCell(colnum);

            if (val is long || val is int || val is decimal || val is short)
            {
                cell.SetCellValue(Convert.ToDouble(val));
            } else
            if (val is TimeSpan ts)
            {
                cell.SetCellValue(ts.ToString());
            } else
            if (val is string s)
            {
                cell.SetCellValue(s);
            } else
            {
                cell.SetCellValue(val.ToString());
            }

            if (colWidth != null) _sheet.SetColumnWidth(colnum, colWidth.Value * 256);
            if (style != null) cell.CellStyle = style;
        }

        #endregion
    }
}
 