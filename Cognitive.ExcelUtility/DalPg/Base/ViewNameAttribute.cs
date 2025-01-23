namespace Cognitive.ExcelUtility.DalPg.Base
{
    public class ViewNameAttribute(string viewName) : Attribute
    {
        public string ViewName { get; private set; } = viewName;
    }
}
