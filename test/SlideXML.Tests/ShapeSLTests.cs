using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using SlideXML.Extensions;
using SlideXML.Models;
using SlideXML.Models.Elements;
using SlideXML.Services;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideXML.Tests
{
    /// <summary>
    /// Represents a unit tests of <see cref="ShapeSL"/> object.
    /// </summary>
    public class ShapeSLTests
    {
        [Fact]
        public void IdHiddenIsPlaceholder_Test()
        {
            var ms = new MemoryStream(Properties.Resources._003);
            var doc = PresentationDocument.Open(ms, false);
            var sldPart = doc.PresentationPart.SlideParts.Single();
            var stubGrFrame = sldPart.Slide.CommonSlideData.ShapeTree.Elements<P.GraphicFrame>().Single(ge => ge.GetId() == 6);

            // ACT
            var shapeBuilder = new ShapeSL.Builder(new BackgroundImageFactory(), new GroupShapeTypeParser(), sldPart);
            var chartShape = shapeBuilder.BuildChartShape(stubGrFrame);

            // CLOSE
            ms.Dispose();
            doc.Dispose();

            // ASSERT
            Assert.Equal(6, chartShape.Id);
            Assert.False(chartShape.Hidden);
            Assert.False(chartShape.IsPlaceholder);
        }

        [Fact]
        public void Hidden_Test()
        {
            var ms = new MemoryStream(Properties.Resources._004);
            var pre = new PresentationSL(ms);

            // ACT
            var allElements = pre.Slides.Single().Shapes;
            var shapeHiddenValue = allElements[0].Hidden;
            var tableHiddenValue = allElements[1].Hidden;

            // CLOSE
            pre.Close();

            // ASSERT
            Assert.True(shapeHiddenValue);
            Assert.False(tableHiddenValue);
        }
    }
}
