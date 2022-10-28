using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Extensions;

internal static class OpenXmlPartExtensions
{
    internal static string AddImagePart(this OpenXmlPart openXmlPart, Stream stream)
    {
        var rIdList = new List<long>();
        var relationships = openXmlPart.ExternalRelationships.Select(r => r.Id)
            .Union(openXmlPart.HyperlinkRelationships.Select(r => r.Id))
            .Union(openXmlPart.DataPartReferenceRelationships.Select(r => r.Id))
            .Union(openXmlPart.Parts.Select(p => p.RelationshipId));
        foreach (var relationship in relationships)
        {
            var match = Regex.Match(relationship, @"\d+");
            if (match.Success)
            {
                var id = long.Parse(match.Value, NumberStyles.None, NumberFormatInfo.CurrentInfo);
                rIdList.Add(id);
            }
        }

        var nextId = 1L;
        if (rIdList.Any())
        {
            nextId = rIdList.Max() + 1;
        }

        var rId = $"rId{nextId}";
        var imagePart = openXmlPart.AddNewPart<ImagePart>("image/png", rId);
        stream.Position = 0;
        imagePart.FeedData(stream);

        return rId;
    }
}