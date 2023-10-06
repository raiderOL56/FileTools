using System.Data;

namespace FileTools.Extensions
{
    public static class DataTableExtensions
    {
        public static bool IsNullOrEmpty(this DataTable dt) => dt == null || dt.Rows.Count == 0;
    }
}