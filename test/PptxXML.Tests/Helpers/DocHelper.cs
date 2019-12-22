using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace PptxXML.Tests.Helpers
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
