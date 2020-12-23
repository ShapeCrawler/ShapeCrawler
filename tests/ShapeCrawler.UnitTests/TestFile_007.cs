using System.IO;
using System.Linq;
using ShapeCrawler.Models;
using SlideDotNet.Models;
using Xunit;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.UnitTests
{
    public class TestFile_007
    {
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
            pre.Close();

            var pre2 = new PresentationEx(ms);
            var numSlides = pre2.Slides.Count();
            var numElements = pre2.Slides.Single().Shapes.Count;
            pre2.Close();
            ms.Dispose();

            // ASSERT
            Assert.Equal(1, numSlides);
            Assert.Equal(2, numElements);
        }

        /// <State>
        /// - there is presentation with two slides;
        /// - first slide is deleted.
        /// </State>
        /// <ExpectedBahavior>
        /// Presentation contains single slide with 1 number.
        /// </ExpectedBahavior>
        [Fact]
        public void Remove_Test1()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._007_2_slides);
            var slides = pre.Slides;
            var slide1 = slides[0];
            var slide2 = slides[1];

            // ACT
            var num1BeforeRemoving = slide1.Number;
            var num2BeforeRemoving = slide2.Number;
            slides.Remove(slide1);
            var num2AfterRemoving = slide2.Number;

            // ARRANGE
            Assert.Equal(1, num1BeforeRemoving);
            Assert.Equal(2, num2BeforeRemoving);
            Assert.Equal(1, num2AfterRemoving);

            // CLEAN
            pre.Close();
        }
    }
}
