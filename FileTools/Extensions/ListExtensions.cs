namespace FileTools.Extensions
{
    public static class ListExtensions
    {
        public static bool IsNullOrEmpty<T>(this List<T> data) => data == null || !data.Any();
        public static bool HasNullItems<T>(this List<T> data) => data != null && data.Any(item => item == null);
    }
}