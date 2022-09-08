using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Extensions
{
    internal static class OpenXmlPartExtensions
    {
        internal static string AddImagePart(this OpenXmlPart openXmlPart, Stream stream)
        {
            var rId = $"rId{Guid.NewGuid().ToString().Replace("-", string.Empty).Substring(0, 5)}";
            var imagePart = openXmlPart.AddNewPart<ImagePart>("image/png", rId);
            imagePart.FeedData(stream);

            return rId;
        }
    }
}