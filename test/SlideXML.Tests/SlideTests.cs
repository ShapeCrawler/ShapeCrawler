using System.IO;
using System.Linq;
using SlideXML.Models;
using Xunit;

namespace SlideXML.Tests
{
    /// <summary>
    /// Represents a unit tests of <see cref="SlideSL"/> object.
    /// </summary>
    public class SlideTests
    {
        [Fact]
        public void ElementsNumber()
        {
            var ms = new MemoryStream(Properties.Resources._002);
            var pre = new PresentationSL(ms);
            var allElements = pre.Slides.First().Shapes;

            // ACT
            var elementsNumber = allElements.Count;

            // CLOSE
            pre.Close();
            
            // ASSERT
            Assert.Equal(3, elementsNumber);
        }
    }
}
