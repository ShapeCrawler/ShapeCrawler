using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using PptxXML.Entities;
using PptxXML.Services;
using Xunit;

namespace PptxXML.Tests
{
    /// <summary>
    /// Represent a unit tests of <see cref="SlideLayoutPartParser"/> object.
    /// </summary>
    public class SlideLayoutPartParserTests
    {
        /// <summary>
        /// Test contains data for title placeholder
        /// </summary>
        [Fact]
        public void GetPlaceholderDataTest()
        {
            var ms = new MemoryStream(Properties.Resources._006_1_slides);
            var xmlDoc = PresentationDocument.Open(ms, false);
            var sldLtPart = xmlDoc.PresentationPart.SlideParts.First().SlideLayoutPart;
            var parser = new SlideLayoutPartParser();

            // ACT
            var phDataDic = parser.GetPlaceholderDic(sldLtPart);

            // CLOSE
            xmlDoc.Close();

            // ASSERT
            Assert.True(phDataDic.Any(d => d.Key.Equals(0)));
        }

        [Fact]
        public void Hidden()
        {
            var ms = new MemoryStream(Properties.Resources._004);
            var pre = new PresentationEx(ms);

            // ACT
            var allElements = pre.Slides.Single().Elements;
            var shapeHiddenValue = allElements[0].Hidden;
            var tableHiddenValue = allElements[1].Hidden;

            // CLOSE
            pre.Dispose();

            // ASSERT
            Assert.True(shapeHiddenValue);
            Assert.False(tableHiddenValue);
        }
    }
}
