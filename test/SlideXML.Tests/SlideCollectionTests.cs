using System.Linq;
using NSubstitute;
using SlideXML.Models;
using SlideXML.Models.Settings;
using SlideXML.Services;
using SlideXML.Services.Placeholders;
using SlideXML.Tests.Helpers;
using Xunit;

namespace SlideXML.Tests
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
            var mockPhService = Substitute.For<IPlaceholderService>();
            var treeParser = new GroupShapeTypeParser();
            var bgImgFactory = new BackgroundImageFactory();
            var mockPreSettings = Substitute.For<IPreSettings>();
            var sldPart = xmlDoc.PresentationPart.SlideParts.First();
            var elementCreator = new ElementFactory(sldPart);

            var newSlide = new SlideSL(sldPart, 1, treeParser, bgImgFactory, mockPreSettings);

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
            var pre = new PresentationSL(Properties.Resources._007_2_slides);
            var slides = pre.Slides;
            var slide1 = slides.First();

            // ACT
            slides.Remove(slide1);

            // ARRANGE
            Assert.Equal(1, slides.Single().Number);

            // CLEAN
            pre.Close();
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
            var pre = new PresentationSL(Properties.Resources._006_1_slides);
            var slides = pre.Slides;
            var slide1 = slides.First();

            // ACT
            slides.Remove(slide1);

            // ARRANGE
            Assert.Empty(slides);

            // CLEAN
            pre.Close();
        }
    }
}
