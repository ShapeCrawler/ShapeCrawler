using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using NSubstitute;
using PptxXML.Extensions;
using PptxXML.Models.Elements;
using PptxXML.Models.Settings;
using PptxXML.Services;
using PptxXML.Services.Builders;
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
            var sldPart = doc.PresentationPart.GetSlidePartByNumber(1);
            var groupShape = sldPart.Slide.CommonSlideData.ShapeTree.Elements<P.GroupShape>().Single(x => x.GetId() == 2);
            var parser = new GroupShapeTypeParser();
            var elFactory = new ElementFactory(new ShapeEx.Builder(new BackgroundImageFactory()));
            var builder = new GroupEx.Builder(parser, elFactory);
            var mockPreSettings = Substitute.For<IPreSettings>();

            // ACT
            var groupEx = builder.Build(groupShape, sldPart, mockPreSettings);

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
