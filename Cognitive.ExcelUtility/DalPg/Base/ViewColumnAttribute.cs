namespace Cognitive.ExcelUtility.DalPg.Base
{
    public class ViewColumnAttribute : Attribute
    {
        public ViewColumnAttribute(string columnName)
        {
            ColumnName = columnName;
        }

        public string ColumnName { get; private set; }
    }
}
