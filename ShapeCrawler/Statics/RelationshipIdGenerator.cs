using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Statics
{
    internal static class RelationshipIdGenerator
    {
        internal static string ForSlidePart(SlidePart slidePart)
        {
            var idList = new List<int>();
            foreach (var idPartPair in slidePart.Parts)
            {
                var matched = Regex.Match(idPartPair.RelationshipId, @"(?<=rId)\d+");
                var hasInt = int.TryParse(matched.Value, out var rIdInt);
                if (hasInt)
                {
                    idList.Add(rIdInt);
                }
            }

            var rId = $"rId{idList.Max() + 1}";

            return rId;
        }
    }
}