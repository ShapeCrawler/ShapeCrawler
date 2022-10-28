using System;

namespace ShapeCrawler.Statics
{
    internal static class RelatedIdGenerator
    {
        public static string Generate()
        {
            return $"rId-{Guid.NewGuid().ToString("N").Substring(0, 5)}";
        }
    }
}