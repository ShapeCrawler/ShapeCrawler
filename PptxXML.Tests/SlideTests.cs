using System.IO;
using System.Linq;
using PptxXML.Models;
using Xunit;

namespace PptxXML.Tests
{
    /// <summary>
    /// Represents a unit tests of <see cref="SlideEx"/> object.
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

            // CLOSE
            pre.Dispose();
            ms.Dispose();
            
            // ASSERT
            Assert.Equal(3, elementsNumber);
        }
    }
}
