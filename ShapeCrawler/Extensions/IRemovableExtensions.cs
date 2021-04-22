namespace ShapeCrawler.Extensions
{
    internal static class IRemovableExtensions
    {
        public static bool IsRemoved(this IRemovable removable)
        {
            return removable.IsRemoved;
        }
    }
}
