namespace Cognitive.ExcelUtility.DalPg.Base
{
    public class ViewColumnAttribute(string columnName) : Attribute
    {
        public string ColumnName { get; private set; } = columnName;
    }
}
