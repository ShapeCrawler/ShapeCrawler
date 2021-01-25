using System.IO;
using System.Linq;
using Xunit;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit
{
    public class TestFile_009 : IClassFixture<TestFile_009Fixture>
    {
        private readonly TestFile_009Fixture _fixture;

        public TestFile_009(TestFile_009Fixture fixture)
        {
            _fixture = fixture;
        }


        [Fact]
        public void Shape_Text_Tests()
        {
            // ARRANGE
            var pre = PresentationSc.Open(Properties.Resources._011_dt, false);
            var grShape = pre.Slides[0].Shapes.Single(s => s.Id == 4);

            // ACT
            var hasTextFrame = grShape.HasTextBox;

            pre.Close();

            // ASSERT
            Assert.False(hasTextFrame);
        }

        [Fact]
        public void Hidden_Test()
        {
            var pre = PresentationSc.Open(Properties.Resources._004, false);

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