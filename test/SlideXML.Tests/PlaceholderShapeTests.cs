using System.IO;
using System.Linq;
using SlideXML.Enums;
using SlideXML.Models;
using SlideXML.Models.Elements;
using Xunit;

namespace SlideXML.Tests
{
    /// <summary>
    /// Contains tests for placeholder shapes.
    /// </summary>
    public class PlaceholderShapeTest
    {
        [Fact]
        public void DateTimePlaceholder_Test()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._008);
            var sp3 = pre.Slides[0].Shapes.Single(sp => sp.Id == 3);

            // ACT
            var hasTextBody = sp3.HasTextBody;

            pre.Dispose();

            // ASSERT
            Assert.False(hasTextBody);
        }
    }
}
