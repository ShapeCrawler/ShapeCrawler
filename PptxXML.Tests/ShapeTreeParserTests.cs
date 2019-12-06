using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using PptxXML.Enums;
using PptxXML.Services;
using Xunit;

namespace PptxXML.Tests
{
    /// <summary>
    /// Contains unit tests of the <see cref="ShapeTreeParser"/> class.
    /// </summary>
    public class ShapeTreeParserTests
    {
        [Fact]
        public void CreateCandidates_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._003);
            var doc = PresentationDocument.Open(ms, false);
            var parser = new ShapeTreeParser();
            var shapeTree = doc.PresentationPart.SlideParts.First().Slide.CommonSlideData.ShapeTree;

            // ACT
            var candidates = parser.CreateCandidates(shapeTree);

            // CLEAN
            doc.Dispose();
            ms.Dispose();

            // ASSERT
            Assert.Single(candidates.Where(c => c.ElementType.Equals(ElementType.Shape)));
            Assert.Single(candidates.Where(c => c.ElementType.Equals(ElementType.Picture)));
            Assert.Single(candidates.Where(c => c.ElementType.Equals(ElementType.Table)));
            Assert.Single(candidates.Where(c => c.ElementType.Equals(ElementType.Chart)));
        }
    }
}
