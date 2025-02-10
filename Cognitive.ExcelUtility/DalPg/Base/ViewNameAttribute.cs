namespace Cognitive.ExcelUtility.DalPg.Base
{
    public class ViewNameAttribute : Attribute
    {
        public ViewNameAttribute(string viewName)
        {
            ViewName = viewName;
        }

        public string ViewName { get; private set; }
    }
}
