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
        public void DateTimePlaceholder_HasText_Test()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._008);
            var sp3 = pre.Slides[0].Shapes.Single(sp => sp.Id == 3);

            // ACT
            var hasTextBody = sp3.HasTextBody;

            pre.Close();

            // ASSERT
            Assert.False(hasTextBody);
        }

        [Fact]
        public void DateTimePlaceholder_Text_Test()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._011_dt);
            var dt = pre.Slides[0].Shapes.Single();

            // ACT
            var text = dt.TextBody.Text;
            var hasText = dt.HasTextBody;

            pre.Close();

            // ASSERT
            Assert.True(hasText);
            Assert.NotNull(text);
        }
    }
}
