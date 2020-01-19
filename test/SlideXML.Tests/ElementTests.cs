using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using SlideXML.Extensions;
using SlideXML.Models;
using SlideXML.Models.Elements;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideXML.Tests
{
    /// <summary>
    /// Represents a unit tests of <see cref="Element"/> object.
    /// </summary>
    public class ElementTests
    {
        [Fact]
        public void IdHiddenIsPlaceholder_Test()
        {
            var ms = new MemoryStream(Properties.Resources._003);
            var doc = PresentationDocument.Open(ms, false);
            var sldPart = doc.PresentationPart.SlideParts.Single();
            var stubGrFrame = sldPart.Slide.CommonSlideData.ShapeTree.Elements<P.GraphicFrame>().Single(ge => ge.GetId() == 6);

            // ACT
            var chart = new ChartEx(stubGrFrame, sldPart);

            // CLOSE
            ms.Dispose();
            doc.Dispose();

            // ASSERT
            Assert.Equal(6, chart.Id);
            Assert.False(chart.Hidden);
            Assert.False(chart.IsPlaceholder);
        }

        [Fact]
        public void Hidden_Test()
        {
            var ms = new MemoryStream(Properties.Resources._004);
            var pre = new PresentationEx(ms);

            // ACT
            var allElements = pre.Slides.Single().Elements;
            var shapeHiddenValue = allElements[0].Hidden;
            var tableHiddenValue = allElements[1].Hidden;

            // CLOSE
            pre.Dispose();
            ms.Dispose();

            // ASSERT
            Assert.True(shapeHiddenValue);
            Assert.False(tableHiddenValue);
        }
    }
}
