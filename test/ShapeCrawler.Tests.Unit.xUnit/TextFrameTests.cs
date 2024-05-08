using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Helpers.Attributes;
using Xunit;

// ReSharper disable All
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit.xUnit
{
    [SuppressMessage("Usage", "xUnit1013:Public method should be marked as test")]
    public class TextFrameTests : SCTest
    {

        

        [Xunit.Theory]
        [MemberData(nameof(TestCasesTextFrameXPath))]
        public void GetPresentationSlideTextFrameXPath(string presentationName, int slideNumber, string[] expectedXPath)
        {
            // Arrange
            var pres = new Presentation(StreamOf(presentationName));
            var slide = pres.Slides[slideNumber];
            var textFrames = slide.TextFrames();

            // Act
            var actualXPath = textFrames.Select(tf => tf.SDKXPath).ToArray();

            // Assert
            actualXPath.Should().BeEquivalentTo(expectedXPath);
        }

        public static IEnumerable<object[]> TestCasesTextFrameXPath()
        {
            yield return new object[]
            {
                "054_get_shape_xpath.pptx", 0,
                new string[]
                {
                    "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[1]/p:txBody[1]",
                    "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[2]/p:txBody[1]"
                }
            };
            yield return new object[]
            {
                "054_get_shape_xpath.pptx", 1,
                new string[]
                {
                    "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[1]/p:txBody[1]",
                    "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[2]/p:txBody[1]",
                    "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[3]/p:txBody[1]"
                }
            };
        }
    }
}