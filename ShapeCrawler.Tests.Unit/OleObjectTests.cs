using System.IO;
using System.Linq;
using ShapeCrawler.Enums;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;

namespace ShapeCrawler.Tests.Unit
{
    public class OleObjectTests : IClassFixture<PptxFixture>
    {
        private readonly PptxFixture _fixture;

        public OleObjectTests(PptxFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void OleObjects_ParseTest()
        {
            // ARRANGE
            var pre = _fixture.Pre009;
            var shapes = pre.Slides[1].Shapes;

            // ACT
            var oleNumbers = shapes.Count(e => e.ContentType.Equals(ShapeContentType.OLEObject));
            var ole9 = shapes.Single(s => s.Id == 9);

            // ASSERT
            Assert.Equal(2, oleNumbers);
            Assert.Equal(485775, ole9.Width);
            Assert.Equal(373062, ole9.Height);
        }
    }
}
