using PptxXML.Entities;
using System.IO;
using System.Linq;
using PptxXML.Entities.Elements;
using PptxXML.Models.Elements;
using Xunit;

namespace PptxXML.Tests
{
    /// <summary>
    /// Represents unit tests of the <see cref="PresentationEx"/> class.
    /// </summary>
    public class PresentationExTests
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

        [Fact]
        public void ShapeTextBody_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._008);

            // ACT
            var shapes = pre.Slides.Single().Elements.OfType<ShapeEx>();
            var sh36 = shapes.Single(e => e.Id == 36);
            var sh37 = shapes.Single(e => e.Id == 37);
            pre.Dispose();

            // ASSERT
            Assert.Null(sh36.TextBody);
            Assert.NotNull(sh37.TextBody);
            Assert.Equal("P1t1 P1t2", sh37.TextBody.Paragraphs[0].Text);
            Assert.Equal("p2", sh37.TextBody.Paragraphs[1].Text);
        }
    }
}
