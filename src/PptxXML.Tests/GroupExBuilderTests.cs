using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using PptxXML.Enums;
using PptxXML.Extensions;
using PptxXML.Models.Elements;
using PptxXML.Services;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxXML.Tests
{
    /// <summary>
    /// Contains tests for the <see cref="GroupEx.Builder"/> class.
    /// </summary>
    public class GroupExBuilderTests
    {
        [Fact]
        public void Build_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._009);
            var doc = PresentationDocument.Open(ms, false);
            var groupShape = doc.PresentationPart.SlideParts.Single().Slide.CommonSlideData.ShapeTree.Elements<P.GroupShape>().Single(x => x.GetId() == 2);
            var parser = new GroupShapeTypeParser();
            var elFactory = new ElementFactory(parser);
            var builder = new GroupEx.Builder(parser, elFactory);

            // ACT
            var groupEx = builder.Build(groupShape);

            // CLEAN
            doc.Dispose();
            ms.Dispose();

            // ASSERT
            Assert.Equal(2, groupEx.Elements.Count);
            Assert.Equal(7547759, groupEx.X);
            Assert.Equal(2372475, groupEx.Y);
            Assert.Equal(1143000, groupEx.Width);
            Assert.Equal(1044645, groupEx.Height);
        }
    }
}
