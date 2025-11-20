using System.IO;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Presentations;

namespace ShapeCrawler.Extensions;

internal static class TypedOpenXmlPartExtensions
{
    internal static string AddImagePart(this OpenXmlPart openXmlPart, Stream stream, string mimeType)
    {
        var rId = new SCOpenXmlPart(openXmlPart).NextRelationshipId();

        var imagePart = openXmlPart.AddNewPart<ImagePart>(mimeType, rId);
        stream.Position = 0;
        imagePart.FeedData(stream);

        return rId;
    }
}