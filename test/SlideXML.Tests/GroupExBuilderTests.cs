using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using NSubstitute;
using SlideXML.Extensions;
using SlideXML.Models.Settings;
using SlideXML.Models.SlideComponents;
using SlideXML.Services;
using SlideXML.Services.Placeholders;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideXML.Tests
{
    /// <summary>
    /// Contains tests for the <see cref="GroupSL.Builder"/> class.
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
            var mockPhService = Substitute.For<IPlaceholderService>();
            var elFactory = new ElementFactory(sldPart);
            var builder = new ShapeSL.Builder(new BackgroundImageFactory(), new GroupShapeTypeParser(), sldPart);
            var mockPreSettings = Substitute.For<IPreSettings>();

            // ACT
            var groupEx = builder.BuildGroup(elFactory, groupShape, mockPreSettings);

            // CLEAN
            doc.Dispose();
            ms.Dispose();

            // ASSERT
            Assert.Equal(2, groupEx.Group.Shapes.Count);
            Assert.Equal(7547759, groupEx.X);
            Assert.Equal(2372475, groupEx.Y);
            Assert.Equal(1143000, groupEx.Width);
            Assert.Equal(1044645, groupEx.Height);
        }
    }
}
