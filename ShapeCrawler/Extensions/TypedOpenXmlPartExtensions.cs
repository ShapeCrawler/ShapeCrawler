using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Extensions;

internal static class TypedOpenXmlPartExtensions
{
    internal static string GetNextRelationshipId(this TypedOpenXmlPart typedOpenXmlPart)
    {
        var idNums = new List<long>();
        var relationships = typedOpenXmlPart.ExternalRelationships.Select(r => r.Id)
            .Union(typedOpenXmlPart.HyperlinkRelationships.Select(r => r.Id))
            .Union(typedOpenXmlPart.DataPartReferenceRelationships.Select(r => r.Id))
            .Union(typedOpenXmlPart.Parts.Select(p => p.RelationshipId));
        foreach (var relationship in relationships)
        {
            var match = Regex.Match(relationship, @"\d+");
            if (match.Success)
            {
                var id = long.Parse(match.Value, NumberStyles.None, NumberFormatInfo.CurrentInfo);
                idNums.Add(id);
            }
        }

        var nextId = 1L;
        if (idNums.Any())
        {
            nextId = idNums.Max() + 1;
        }
        
        return $"rId{nextId}";        
    }
    
    internal static string AddImagePart(this TypedOpenXmlPart typedOpenXmlPart, Stream stream)
    {
        var rId = typedOpenXmlPart.GetNextRelationshipId();
        
        var imagePart = typedOpenXmlPart.AddNewPart<ImagePart>("image/png", rId);
        stream.Position = 0;
        imagePart.FeedData(stream);

        return rId;
    }
}