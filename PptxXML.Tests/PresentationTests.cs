using DocumentFormat.OpenXml.Presentation;
using PptxXML.Entities;
using System.IO;
using System.Linq;
using Xunit;

namespace PptxXML.Tests
{
    /// <summary>
    /// Represent unit tests of <see cref="Presentation"/> object
    /// </summary>
    public class PresentationTests
    {
        [Fact]
        public void SlidesNumber_Test()
        {
            var ms = new MemoryStream(Properties.Resources._001);
            var pre = new PresentationEx(ms);

            // ACT
            var sldNumber = pre.Slides.Count();

            // CLOSE
            pre.Dispose();
            
            // ASSERT
            Assert.Equal(2, sldNumber);
        }

        /// <State>
        /// - there is a presentation with two slides. The first slide contains one element. The second slide includes two elements;
        /// - the first slide is removed;
        /// - the presentation is closed.
        /// </State>
        /// <ExpectedBahavior>
        /// There is the only second slide with its two elements.
        /// </ExpectedBahavior>
        [Fact]
        public void SlidesRemove_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._007_2_slides);
            var pre = new PresentationEx(ms);

            // ACT
            var slide1 = pre.Slides.First();
            pre.Slides.Remove(slide1);
            pre.Dispose();

            var pre2 = new PresentationEx(ms);
            var numSlides = pre2.Slides.Count();
            var numElements = pre2.Slides.Single().Elements.Count;
            pre2.Dispose();
            ms.Dispose();

            // ASSERT
            Assert.Equal(1, numSlides);
            Assert.Equal(2, numElements);
        }
    }
}
