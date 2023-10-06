namespace FileTools.Models.Xlsx
{
    public class CustomHeader
    {
        public string PropertyName { get; }
        public string HeaderName { get; }
        public CustomStyle? CustomStyle { get; }
        public CustomHeader(string PropertyName, string HeaderName)
        {
            this.PropertyName = PropertyName;
            this.HeaderName = HeaderName;
        }
        public CustomHeader(string PropertyName, string HeaderName, CustomStyle CustomStyle)
        {
            this.PropertyName = PropertyName;
            this.HeaderName = HeaderName;
            this.CustomStyle = CustomStyle;
        }
    }
}