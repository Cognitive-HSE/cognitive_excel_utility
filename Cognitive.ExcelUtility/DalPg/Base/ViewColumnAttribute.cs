namespace Cognitive.ExcelUtility.DalPg.Base
{
    public class ViewColumnAttribute : Attribute
    {
        public string ColumnName { get; private set; }

        public ViewColumnAttribute(string columnName)
        {
            ColumnName = columnName;
        }
    }
}
