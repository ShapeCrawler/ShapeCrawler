using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace SlideXML.Tests.Helpers
{
    /// <summary>
    /// Represents a helper for <see cref="PresentationDocument"/> class.
    /// </summary>
    public static class DocHelper
    {
        public static PresentationDocument Open(byte[] fileBytes)
        {
            var stream = new MemoryStream(fileBytes);

            return PresentationDocument.Open(stream, false);
        }
    }
}
