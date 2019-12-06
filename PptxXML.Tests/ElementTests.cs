using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using PptxXML.Enums;
using PptxXML.Extensions;
using PptxXML.Models;
using PptxXML.Models.Elements;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxXML.Tests
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
            var stubGrFrame = doc.PresentationPart.SlideParts.Single().Slide.CommonSlideData.ShapeTree.Elements<P.GraphicFrame>().Single(ge => ge.GetId() == 4);

            // ACT
            var chart = new ChartEx(stubGrFrame);

            // CLOSE
            ms.Dispose();
            doc.Dispose();

            // ASSERT
            Assert.Equal(4, chart.Id);
            Assert.False(chart.Hidden);
            Assert.False(chart.IsPlaceholder);
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
            ms.Dispose();

            // ASSERT
            Assert.True(shapeHiddenValue);
            Assert.False(tableHiddenValue);
        }
    }
}
