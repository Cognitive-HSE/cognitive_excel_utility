using Cognitive.ExcelUtility.DalPg.Entities;
using Cognitive.ExcelUtility.DalPg.Managers;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Cognitive.ExcelUtility
{
    internal static class ExcelBuilder
    {
        private static readonly XSSFWorkbook Workbook = new();
        private static readonly ISheet Sheet = Workbook.CreateSheet("Report");

        private const int HeaderRows = 2;
        private const int TestInfoCols = 5;
        private const int UserInfoCols = 15;

        #region User Readable Report

        public static async Task<IWorkbook> CreateUserReadableReport()
        {
            var headerNames = await GetHeaderNames();
            var overallCols = (headerNames.Count * TestInfoCols) + UserInfoCols;
            
            IRow[] headerRows = [Sheet.CreateRow(0), Sheet.CreateRow(1)];
            Sheet.CreateFreezePane(UserInfoCols, HeaderRows);

            // Сборка шапки юзеров
            headerRows[0].CreateAndWriteCell(0, "ИД", CellStyleHeader, 5);
            headerRows[0].CreateAndWriteCell(1, "Возраст", CellStyleHeader, 9);
            headerRows[0].CreateAndWriteCell(2, "Образование", CellStyleHeader, 20);
            headerRows[0].CreateAndWriteCell(3, "Специальность", CellStyleHeader, 20);
            headerRows[0].CreateAndWriteCell(4, "Место жительства", CellStyleHeader, 20);
            headerRows[0].CreateAndWriteCell(5, "Рост", CellStyleHeader, 5);
            headerRows[0].CreateAndWriteCell(6, "Вес", CellStyleHeader, 5);
            headerRows[0].CreateAndWriteCell(7, "Ведущая рука", CellStyleHeader, 10);
            headerRows[0].CreateAndWriteCell(8, "Заболевания", CellStyleHeader, 15);
            headerRows[0].CreateAndWriteCell(9, "Курение", CellStyleHeader, 9);
            headerRows[0].CreateAndWriteCell(10, "Алкоголь", CellStyleHeader, 10);
            headerRows[0].CreateAndWriteCell(11, "Спорт", CellStyleHeader, 15);
            headerRows[0].CreateAndWriteCell(12, "Бессонница", CellStyleHeader, 7);
            headerRows[0].CreateAndWriteCell(13, "Текущее самочувствие", CellStyleHeader, 15);
            headerRows[0].CreateAndWriteCell(14, "Геймер", CellStyleHeader, 8);


            for (var i = 0; i < UserInfoCols; i++)
                Sheet.AddMergedRegion(new CellRangeAddress(0, 1, i, i));

            // Сборка шапки тестов
            for (var i = 0; i < headerNames.Count; i++)
            {
                var startRegionCol = UserInfoCols + i * TestInfoCols;
                headerNames[i].StartRegionCol = startRegionCol;
                headerRows[0].CreateAndWriteCell(startRegionCol, headerNames[i].TestName, CellStyleHeader);
                Sheet.AddMergedRegion(new CellRangeAddress(0, 0, startRegionCol, startRegionCol + TestInfoCols - 1));

                headerRows[1].CreateAndWriteCell(startRegionCol + 0, "#", CellStyleHeader, 3);
                headerRows[1].CreateAndWriteCell(startRegionCol + 1, "% выполнения", CellStyleHeader, 12);
                headerRows[1].CreateAndWriteCell(startRegionCol + 2, "Время выполнения", CellStyleHeader, 12);
                headerRows[1].CreateAndWriteCell(startRegionCol + 3, headerNames[i].TitleCorrect, CellStyleHeader, 20);
                headerRows[1].CreateAndWriteCell(startRegionCol + 4, headerNames[i].TitleAll, CellStyleHeader, 20);
            }

            // Наполнение таблицы
            var data = await GetTestResults();
            var userData = await GetUsers();
            var rownum = 2; // 2 == высота шапки

            foreach (var user in data.GroupBy(el => el.UserId))
            {
                var rowsCreated = 0;
                foreach (var testTry in user.GroupBy(el => el.TryNumber))
                {
                    var row = Sheet.CreateRow(rownum++);
                    rowsCreated++;
                    if (rowsCreated == 1)
                    {
                        row.CreateAndWriteCell(0, user.Key, CellStyleData);
                        WriteUserData(row, 0, ref userData, user.Key);
                    }

                    foreach (var testRes in testTry)
                        WriteTestResult(row, headerNames.First(el => el.TestId == testRes.TestId).StartRegionCol, testRes);
                }
                
                if (rowsCreated > 1)
                {
                    for (var i = 0; i < UserInfoCols; i++)
                        Sheet.AddMergedRegion(new CellRangeAddress(rownum - rowsCreated, rownum - 1, i, i));
                }

                // Рисуем горизонтальную границу между людьми
                for (var i = 0; i < overallCols; i++)
                {
                    var row = Sheet.GetRow(rownum - 1);
                    var cell = row.GetOrCreateCell(i);
                    cell.DrawBorder(false);
                }
            }

            // Рисуем вертикальные границы
            for (var i = 0; i < rownum; i++)
            {
                var row = Sheet.GetRow(i);

                for (var j = UserInfoCols - 1; j < overallCols; j += TestInfoCols)
                {
                    var cell = row.GetOrCreateCell(j);
                    cell.DrawBorder(true);
                }
            }

            return Workbook;
        }

        private static void WriteTestResult(IRow row, int startCol, TestResultEntity item)
        {
            row.CreateAndWriteCell(startCol + 0, item.TryNumber, CellStyleData);
            row.CreateAndWriteCell(startCol + 1, item.Accuracy, CellStyleData);
            row.CreateAndWriteCell(startCol + 2, item.CompleteTime, CellStyleData);
            row.CreateAndWriteCell(startCol + 3, item.NumberCorrectAnswers, CellStyleData);
            row.CreateAndWriteCell(startCol + 4, item.NumberAllAnswers, CellStyleData);
        }

        private static void WriteUserData(IRow row, int startCol, ref List<UserEntity> users, long? userId)
        {
            var item = users.FirstOrDefault(el => el.UserId == userId);
            if (item == null)
            {
                row.CreateAndWriteCell(startCol + 0, userId, CellStyleData);
                return;
            }

            row.CreateAndWriteCell(startCol + 0, item.UserId, CellStyleData);
            row.CreateAndWriteCell(startCol + 1, item.Age, CellStyleData);
            row.CreateAndWriteCell(startCol + 2, item.Education, CellStyleData);
            row.CreateAndWriteCell(startCol + 3, item.Speciality, CellStyleData);
            row.CreateAndWriteCell(startCol + 4, item.Residence, CellStyleData);
            row.CreateAndWriteCell(startCol + 5, item.Height, CellStyleData);
            row.CreateAndWriteCell(startCol + 6, item.Weight, CellStyleData);
            row.CreateAndWriteCell(startCol + 7, item.LeadHand, CellStyleData);
            row.CreateAndWriteCell(startCol + 8, item.Diseases, CellStyleData);
            row.CreateAndWriteCell(startCol + 9, item.Smoking, CellStyleData);
            row.CreateAndWriteCell(startCol + 10, item.Alcohol, CellStyleData);
            row.CreateAndWriteCell(startCol + 11, item.Sport, CellStyleData);
            row.CreateAndWriteCell(startCol + 12, item.Insomnia, CellStyleData);
            row.CreateAndWriteCell(startCol + 13, item.CurrentHealth, CellStyleData);
            row.CreateAndWriteCell(startCol + 14, item.Gaming, CellStyleData);
        }

        #endregion

        #region Styles

        private static readonly ICellStyle CellStyleHeader = CellStyleHeaderBuild();
        private static readonly ICellStyle CellStyleData = CellStyleDataBuild();

        private static ICellStyle CellStyleHeaderBuild()
        {
            var style = Workbook.CreateCellStyle();

            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.WrapText = true;

            return style;
        }

        private static ICellStyle CellStyleDataBuild()
        {
            var style = Workbook.CreateCellStyle();

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

        private static async Task<List<UserEntity>> GetUsers()
        {
            return await new UserManager().ReadAll();
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
            switch (val)
            {
                case null:
                    return;
                case long or int or decimal or short:
                    cell.SetCellValue(Convert.ToDouble(val));
                    break;
                case TimeSpan ts:
                    cell.SetCellValue(ts.ToString());
                    break;
                case string s:
                    cell.SetCellValue(s);
                    break;
                case bool b:
                    cell.SetCellValue(b ? "Да" : "Нет");
                    break;
            }

            if (colWidth != null) Sheet.SetColumnWidth(colnum, colWidth.Value * 256);
            if (style != null) cell.CellStyle = style;
        }

        private static void DrawBorder(this ICell cell, bool right)
        {
            var newstyle = Workbook.CreateCellStyle();
            newstyle.CloneStyleFrom(cell.CellStyle);
            
            if (right) 
                newstyle.BorderRight = BorderStyle.Thin;
            else 
                newstyle.BorderBottom = BorderStyle.Thin;

            cell.CellStyle = newstyle;
        }

        private static ICell GetOrCreateCell(this IRow row, int idx)
        {
            var cell = row.GetCell(idx) ?? row.CreateCell(idx);
            return cell;
        }

        #endregion
    }
}
 