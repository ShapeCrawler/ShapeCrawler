using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using SlideXML.Enums;
using SlideXML.Services;
using Xunit;

namespace SlideXML.Tests
{
    /// <summary>
    /// Contains unit tests of the <see cref="GroupShapeTypeParser"/> class.
    /// </summary>
    public class GroupShapeTypeParserTests
    {
        [Fact]
        public void CreateCandidates_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._003);
            var doc = PresentationDocument.Open(ms, false);
            var parser = new GroupShapeTypeParser();
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
            Assert.Single(candidates.Where(c => c.ElementType.Equals(ElementType.Group)));
        }
    }
}
