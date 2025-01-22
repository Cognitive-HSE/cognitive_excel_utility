namespace Cognitive.ExcelUtility.DalPg.Base
{
    public class ViewNameAttribute : Attribute
    {
        public string ViewName { get; private set; }

        public ViewNameAttribute(string viewName)
        {
            ViewName = viewName;
        }
    }
}
