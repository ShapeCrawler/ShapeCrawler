using System.IO;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shared;

namespace ShapeCrawler.Extensions;

internal static class TypedOpenXmlPartExtensions
{
    internal static (string, ImagePart) AddImagePart(this OpenXmlPart typedOpenXmlPart, Stream stream, string mimeType)
    {
        var rId = new SOpenXmlPart(typedOpenXmlPart).NextRelationshipId();

        var imagePart = typedOpenXmlPart.AddNewPart<ImagePart>(mimeType, rId);
        stream.Position = 0;
        imagePart.FeedData(stream);

        return (rId, imagePart);
    }
}