using System.IO;
using System.Linq;
using PptxXML.Entities;
using PptxXML.Enums;
using Xunit;

namespace PptxXML.Tests
{
    /// <summary>
    /// Represent a unit tests of <see cref="Element"/> object.
    /// </summary>
    public class ElementTests
    {
        [Fact]
        public void Type()
        {
            var ms = new MemoryStream(Properties.Resources._003);
            var pre = new PresentationEx(ms);
            var allElements = pre.Slides.First().Elements;

            // ACT
            var chart = allElements[0].Type;
            var group = allElements[1].Type;
            var picture = allElements[2].Type;
            var shape = allElements[3].Type;
            var table = allElements[4].Type;

            // CLOSE
            pre.Dispose();
            
            // ASSERT
            Assert.Equal(ElementType.Chart, chart);
            Assert.Equal(ElementType.Group, group);
            Assert.Equal(ElementType.Picture, picture);
            Assert.Equal(ElementType.Shape, shape);
            Assert.Equal(ElementType.Table, table);
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
