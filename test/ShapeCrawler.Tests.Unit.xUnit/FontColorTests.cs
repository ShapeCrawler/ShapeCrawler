using System.Collections.Generic;
using FluentAssertions;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;

namespace ShapeCrawler.Tests.Unit;

public class FontColorTests : SCTest
{
    [Theory]
    [MemberData(nameof(TestCasesColorGetter))]
    public void ColorHex_Getter_returns_color(TestCase<IParagraph, string> testCase)
    {
        // Arrange
        var paragraph = testCase.Param1;
        var colorHex = testCase.Param2;
        var colorFormat = paragraph.Portions[0].Font.Color;
        var expectedColor = colorHex;

        // Act
        var actualColor = colorFormat.ColorHex;

        // Assert
        actualColor.Should().Be(expectedColor);
    }

    public static IEnumerable<object[]> TestCasesColorGetter
    {
        get
        {
            var stream1 = StreamOf("020.pptx");
            var pres1 = new SCPresentation(stream1);
            var paragraph1 = pres1.Slides[0].Shapes.GetById<IShape>(2).TextFrame!.Paragraphs[0];
            var testCase1 = new TestCase<IParagraph, string>(1, paragraph1, "000000");
            yield return new object[] { testCase1 };

            var stream2 = StreamOf("020.pptx");
            var pres2 = new SCPresentation(stream2);
            var paragraph2 = pres2.Slides[0].Shapes.GetById<IShape>(3).TextFrame!.Paragraphs[0];
            var testCase2 = new TestCase<IParagraph, string>(2, paragraph2, "000000");
            yield return new object[] { testCase2 };

            var stream3 = StreamOf("020.pptx");
            var pres3 = new SCPresentation(stream3);
            var paragraph3 = pres3.Slides[2].Shapes.GetById<IShape>(8).TextFrame!.Paragraphs[1];
            var testCase3 = new TestCase<IParagraph, string>(3, paragraph3, "FFFF00");
            yield return new object[] { testCase3 };

            var stream4 = StreamOf("001.pptx");
            var pres4 = new SCPresentation(stream4);
            var paragraph4 = pres4.Slides[0].Shapes.GetById<IShape>(4).TextFrame!.Paragraphs[0];
            var testCase4 = new TestCase<IParagraph, string>(4, paragraph4, "000000");
            yield return new object[] { testCase4 };

            var stream5 = StreamOf("002.pptx");
            var pres5 = new SCPresentation(stream5);
            var paragraph5 = pres5.Slides[1].Shapes.GetById<IShape>(3).TextFrame!.Paragraphs[0];
            var testCase5 = new TestCase<IParagraph, string>(5, paragraph5, "000000");
            yield return new object[] { testCase5 };

            var stream6 = StreamOf("026.pptx");
            var pres6 = new SCPresentation(stream6);
            var paragraph6 = pres6.Slides[0].Shapes.GetById<IShape>(128).TextFrame!.Paragraphs[0];
            var testCase6 = new TestCase<IParagraph, string>(6, paragraph6, "000000");
            yield return new object[] { testCase6 };

            var stream7 = StreamOf("autoshape-case017_slide-number.pptx");
            var pres7 = new SCPresentation(stream7);
            var paragraph7 = pres7.Slides[0].Shapes.GetById<IShape>(5).TextFrame!.Paragraphs[0];
            var testCase7 = new TestCase<IParagraph, string>(7, paragraph7, "000000");
            yield return new object[] { testCase7 };

            var stream8 = StreamOf("031.pptx");
            var pres8 = new SCPresentation(stream8);
            var paragraph8 = pres8.Slides[0].Shapes.GetById<IShape>(44).TextFrame!.Paragraphs[0];
            var testCase8 = new TestCase<IParagraph, string>(8, paragraph8, "000000");
            yield return new object[] { testCase8 };

            var stream9 = StreamOf("033.pptx");
            var pres9 = new SCPresentation(stream9);
            var paragraph9 = pres9.Slides[0].Shapes.GetById<IShape>(3).TextFrame!.Paragraphs[0];
            var testCase9 = new TestCase<IParagraph, string>(9, paragraph9, "000000");
            yield return new object[] { testCase9 };

            var stream10 = StreamOf("038.pptx");
            var pres10 = new SCPresentation(stream10);
            var paragraph10 = pres10.Slides[0].Shapes.GetById<IShape>(102).TextFrame!.Paragraphs[0];
            var testCase10 = new TestCase<IParagraph, string>(10, paragraph10, "000000");
            yield return new object[] { testCase10 };

            var stream11 = StreamOf("001.pptx");
            var pres11 = new SCPresentation(stream11);
            var paragraph11 = pres11.Slides[2].Shapes.GetById<IShape>(4).TextFrame!.Paragraphs[0];
            var testCase11 = new TestCase<IParagraph, string>(11, paragraph11, "000000");
            yield return new object[] { testCase11 };

            var stream12 = StreamOf("001.pptx");
            var pres12 = new SCPresentation(stream12);
            var paragraph12 = pres12.Slides[4].Shapes.GetById<IShape>(5).TextFrame!.Paragraphs[0];
            var testCase12 = new TestCase<IParagraph, string>(12, paragraph12, "000000");
            yield return new object[] { testCase12 };

            var stream13 = StreamOf("014.pptx");
            var pres13 = new SCPresentation(stream13);
            var paragraph13 = pres13.Slides[0].Shapes.GetById<IShape>(61).TextFrame!.Paragraphs[0];
            var testCase13 = new TestCase<IParagraph, string>(13, paragraph13, "595959");
            yield return new object[] { testCase13 };

            var stream14 = StreamOf("014.pptx");
            var pres14 = new SCPresentation(stream14);
            var paragraph14 = pres14.Slides[5].Shapes.GetById<IShape>(52).TextFrame!.Paragraphs[0];
            var testCase14 = new TestCase<IParagraph, string>(14, paragraph14, "FFFFFF");
            yield return new object[] { testCase14 };

            var stream15 = StreamOf("032.pptx");
            var pres15 = new SCPresentation(stream15);
            var paragraph15 = pres15.Slides[0].Shapes.GetById<IShape>(10242).TextFrame!.Paragraphs[0];
            var testCase15 = new TestCase<IParagraph, string>(15, paragraph15, "0070C0");
            yield return new object[] { testCase15 };

            var stream16 = StreamOf("034.pptx");
            var pres16 = new SCPresentation(stream16);
            var paragraph16 = pres16.Slides[0].Shapes.GetById<IShape>(2).TextFrame!.Paragraphs[0];
            var testCase16 = new TestCase<IParagraph, string>(16, paragraph16, "000000");
            yield return new object[] { testCase16 };

            var stream17 = StreamOf("035.pptx");
            var pres17 = new SCPresentation(stream17);
            var paragraph17 = pres17.Slides[0].Shapes.GetById<IShape>(9).TextFrame!.Paragraphs[0];
            var testCase17 = new TestCase<IParagraph, string>(17, paragraph17, "000000");
            yield return new object[] { testCase17 };

            var stream18 = StreamOf("036.pptx");
            var pres18 = new SCPresentation(stream18);
            var paragraph18 = pres18.Slides[0].Shapes.GetById<IShape>(6146).TextFrame!.Paragraphs[0];
            var testCase18 = new TestCase<IParagraph, string>(18, paragraph18, "404040");
            yield return new object[] { testCase18 };

            var stream19 = StreamOf("037.pptx");
            var pres19 = new SCPresentation(stream19);
            var paragraph19 = pres19.Slides[0].Shapes.GetById<IShape>(7).TextFrame!.Paragraphs[0];
            var testCase19 = new TestCase<IParagraph, string>(19, paragraph19, "1A1A1A");
            yield return new object[] { testCase19 };
        }
    }
}