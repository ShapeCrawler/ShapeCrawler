using PptxXML.Models;
using PptxXML.Tests.Helpers;
using System.Linq;
using PptxXML.Services;
using Xunit;

namespace PptxXML.Tests
{
    /// <summary>
    /// Represents unit tests for the <see cref="SlideCollection"/> class.
    /// </summary>
    public class SlideCollectionTests
    {
        [Fact]
        public void Add_AddedOneItem_SlidesNumberIsOne()
        {
            // ARRANGE
            var xmlDoc = DocHelper.Open(Properties.Resources._001);
            var slides = new SlideCollection(xmlDoc);
            var newSlide = new SlideEx(xmlDoc.PresentationPart.SlideParts.First(), 1, new ElementCreator());

            // ACT
            slides.Add(newSlide);

            // CLEAN
            xmlDoc.Dispose();

            // ASSERT
            Assert.Single(slides);
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
            var slide1 = slides.First();

            // ACT
            slides.Remove(slide1);

            // ARRANGE
            Assert.Equal(1, slides.Single().Number);

            // CLEAN
            pre.Dispose();
        }

        /// <State>
        /// - there is a single slide presentation;
        /// - slide is deleted.
        /// </State>
        /// <ExpectedBahavior>
        /// Presentation is empty.
        /// </ExpectedBahavior>
        [Fact]
        public void Remove_Test2()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._006_1_slides);
            var slides = pre.Slides;
            var slide1 = slides.First();

            // ACT
            slides.Remove(slide1);

            // ARRANGE
            Assert.Empty(slides);

            // CLEAN
            pre.Dispose();
        }
    }
}
