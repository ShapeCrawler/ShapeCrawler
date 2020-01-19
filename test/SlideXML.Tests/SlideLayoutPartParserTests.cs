using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using SlideXML.Services.Placeholders;
using Xunit;

namespace SlideXML.Tests
{
    /// <summary>
    /// Represent a unit tests of <see cref="SlideLayoutPartParser"/> object.
    /// </summary>
    public class SlideLayoutPartParserTests
    {
        /// <summary>
        /// Tests contains data for title placeholder.
        /// </summary>
        [Fact]
        public void GetPlaceholderDic_Test()
        {
            var ms = new MemoryStream(Properties.Resources._006_1_slides);
            var xmlDoc = PresentationDocument.Open(ms, false);
            var sldLtPart = xmlDoc.PresentationPart.SlideParts.First().SlideLayoutPart;
            var parser = new SlideLayoutPartParser();

            // ACT
            var phDataDic = parser.GetPlaceholderDic(sldLtPart);

            // CLOSE
            xmlDoc.Close();
            ms.Dispose();

            // ASSERT
            Assert.Contains(phDataDic, d => d.Key.Equals(100));
        }
    }
}
