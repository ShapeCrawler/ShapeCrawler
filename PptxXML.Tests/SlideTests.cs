using PptxXML.Entities;
using System.IO;
using System.Linq;
using Xunit;

namespace PptxXML.Tests
{
    /// <summary>
    /// Represent a unit tests of <see cref="SlideEx"/> object.
    /// </summary>
    public class SlideTests
    {
        [Fact]
        public void ElementsNumber()
        {
            var ms = new MemoryStream(Properties.Resources._002);
            var pre = new PresentationEx(ms);
            var allElements = pre.Slides.First().Elements;

            // ACT
            var elementsNumber = allElements.Count;
            var firstElementId = allElements[0].Id;

            // CLOSE
            pre.Dispose();
            
            // ASSERT
            Assert.Equal(5, elementsNumber);
            Assert.Equal(2, firstElementId);
        }
    }
}
