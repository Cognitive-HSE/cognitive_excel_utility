using Cognitive.ExcelUtility.DalPg.Entities;
using Cognitive.ExcelUtility.DalPg.Managers;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Cognitive.ExcelUtility
{
    internal static class ExcelBuilder
    {
        private static readonly XSSFWorkbook _workbook = new XSSFWorkbook();
        private readonly static ISheet _sheet = _workbook.CreateSheet("Report");

        #region User Readable Report

        public static async Task<IWorkbook> CreateUserReadableReport()
        {
            const int HeaderRows = 2;
            const int TestInfoCols = 5;
            const int UserInfoCols = 1;

            var headerNames = await GetHeaderNames();
            int overallCols = (headerNames.Count * TestInfoCols) + UserInfoCols;
            
            IRow[] headerRows = { _sheet.CreateRow(0), _sheet.CreateRow(1) };
            _sheet.CreateFreezePane(UserInfoCols, HeaderRows);

            // Сборка шапки юзеров
            headerRows[0].CreateAndWriteCell(0, "ИД пользователя", CellStyleHeader, 15);
            _sheet.AddMergedRegion(new CellRangeAddress(0, 1, 0, 0));

            // Сборка шапки тестов
            for (int i = 0; i < headerNames.Count; i++)
            {
                var startRegionCol = UserInfoCols + i * TestInfoCols;
                headerNames[i].StartRegionCol = startRegionCol;
                headerRows[0].CreateAndWriteCell(startRegionCol, headerNames[i].TestName, CellStyleHeader);
                _sheet.AddMergedRegion(new CellRangeAddress(0, 0, startRegionCol, startRegionCol + TestInfoCols - 1));

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
                        row.CreateAndWriteCell(0, user.Key, CellStyleData);
                    }

                    foreach (var testRes in testTry)
                        WriteTestResult(row, headerNames.First(el => el.TestId == testRes.TestId).StartRegionCol, testRes);
                }
                
                if (rowsCreated > 1)
                {
                    _sheet.AddMergedRegion(new CellRangeAddress(rownum - rowsCreated, rownum - 1, 0, 0));
                }

                // Рисуем горизонтальную границу между людьми
                for (var i = 0; i < overallCols; i++)
                {
                    var row = _sheet.GetRow(rownum - 1);
                    var cell = row.GetOrCreateCell(i);
                    cell.DrawBorder(false);
                }
            }

            // Рисуем вертикальные границы
            for (var i = 0; i < rownum; i++)
            {
                var row = _sheet.GetRow(i);

                for (var j = UserInfoCols - 1; j < overallCols; j += TestInfoCols)
                {
                    var cell = row.GetOrCreateCell(j);
                    cell.DrawBorder(true);
                }
            }

            return _workbook;
        }

        private static void WriteTestResult(IRow row, int startCol, TestResultEntity item)
        {
            row.CreateAndWriteCell(startCol + 0, item.TryNumber, CellStyleData);
            row.CreateAndWriteCell(startCol + 1, item.Accuracy, CellStyleData);
            row.CreateAndWriteCell(startCol + 2, item.CompleteTime, CellStyleData);
            row.CreateAndWriteCell(startCol + 3, item.NumberCorrectAnswers, CellStyleData);
            row.CreateAndWriteCell(startCol + 4, item.NumberAllAnswers, CellStyleData);
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
        private static ICellStyle CellStyleData = CellStyleDataBuild();

        private static ICellStyle CellStyleHeaderBuild()
        {
            var style = _workbook.CreateCellStyle();

            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.WrapText = true;

            return style;
        }

        private static ICellStyle CellStyleDataBuild()
        {
            var style = _workbook.CreateCellStyle();

            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;

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
            var cell = row.CreateCell(colnum);
            if (val == null) return;

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

        private static void DrawBorder(this ICell cell, bool right)
        {
            var newstyle = _workbook.CreateCellStyle();
            newstyle.CloneStyleFrom(cell.CellStyle);
            
            if (right) 
                newstyle.BorderRight = BorderStyle.Thin;
            else 
                newstyle.BorderBottom = BorderStyle.Thin;

            cell.CellStyle = newstyle;
        }

        private static ICell GetOrCreateCell(this IRow row, int idx)
        {
            var cell = row.GetCell(idx);
            if (cell == null)
                cell = row.CreateCell(idx);
            return cell;
        }

        #endregion
    }
}
 