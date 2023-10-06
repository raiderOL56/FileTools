namespace FileTools.Models.Xlsx
{
    public class CellFormat
    {
        public Format Format { get; }
        public string? CustomFormat { get; set; }
        public CellFormat(Format Format)
        {
            this.Format = Format;
        }
        public CellFormat(Format Format, string CustomFormat)
        {
            this.Format = Format;
            this.CustomFormat = CustomFormat;
        }
    }
    public enum Format
    {
        Date,
        Number,
        Decimal
    }
}