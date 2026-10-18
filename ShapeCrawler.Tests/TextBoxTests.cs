using FluentAssertions;
using System.Text.RegularExpressions;
using ShapeCrawler.Tests.Helpers;
using ShapeCrawler.Texts;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tests;

public class TextBoxTests : SCTest
{
    [Test]
    public void HighlightText_spans_runs_preserving_text_formatting_and_hyperlinks()
    {
        // Arrange
        using var pres = ReplacementPresentation(new A.Paragraph(
            new A.ParagraphProperties { Level = 1 },
            new A.Run(new A.RunProperties { Bold = true, FontSize = 1800 }, new A.Text("pre XY")),
            new A.Run(new A.Text(string.Empty)),
            new A.Run(new A.RunProperties { Italic = true, FontSize = 2400 }, new A.Text("Z tail"))));
        var box = pres.Slides[0].Shapes.Last().TextBox!;
        box.Paragraphs[0].Portions.First().Link.AddFile("https://example.com/report");
        var before = box.Text;

        // Act
        var count = box.HighlightText("XYZ", new Color("FFFF00"));

        // Assert
        count.Should().Be(1);
        box.Text.Should().Be(before);
        using var saved = new MemoryStream();
        pres.Save(saved);
        using (var sdk = PresentationDocument.Open(saved, false))
        {
            var paragraph = sdk.PresentationPart!.SlideParts.Single().Slide.Descendants<P.Shape>().Last().TextBody!.Elements<A.Paragraph>().Single();
            var runs = paragraph.Elements<A.Run>().ToArray();
            runs.Select(run => run.Text!.Text).Should().Equal("pre ", "XY", "", "Z", " tail");
            runs.Where(run => run.RunProperties?.GetFirstChild<A.Highlight>() is not null)
                .Select(run => run.Text!.Text).Should().Equal("XY", "Z");
            runs[0].RunProperties!.Bold!.Value.Should().BeTrue();
            runs[1].RunProperties!.FontSize!.Value.Should().Be(1800);
            runs[3].RunProperties!.Italic!.Value.Should().BeTrue();
            runs[4].RunProperties!.FontSize!.Value.Should().Be(2400);
            runs[0].RunProperties!.GetFirstChild<A.HyperlinkOnClick>()!.Id!.Value.Should()
                .Be(runs[1].RunProperties!.GetFirstChild<A.HyperlinkOnClick>()!.Id!.Value);
            sdk.PresentationPart.SlideParts.Single().HyperlinkRelationships.Should().ContainSingle();
            paragraph.ParagraphProperties!.Level!.Value.Should().Be(1);
        }

        using var reopened = new Presentation(saved);
        reopened.Slides[0].Shapes.Last().TextBox.Text.Should().Be(before);
        ValidatePresentation(reopened);
    }

    [TestCase("aaaa", "aa", 2, "aaaa")]
    [TestCase("ababa", "aba", 1, "aba")]
    [TestCase("a <&😀> b", "<&😀>", 1, "<&😀>")]
    [TestCase("abc", "missing", 0, "")]
    public void HighlightText_highlights_original_non_overlapping_matches(string original, string query, int expected, string selected)
    {
        // Arrange
        using var pres = ReplacementPresentation(new A.Paragraph(new A.Run(new A.Text(original))));
        var box = pres.Slides[0].Shapes.Last().TextBox;

        // Act
        var count = box.HighlightText(query, new Color("FFFF00"));

        // Assert
        count.Should().Be(expected);
        box.Text.Should().Be(original);
        string.Concat(box.Paragraphs[0].Portions.Where(portion => !portion.TextHighlightColor.IsTransparent).Select(portion => portion.Text))
            .Should().Be(selected);
        ValidatePresentation(pres);
    }

    [TestCase(false, 1, "file")]
    [TestCase(true, 2, "FILEfile")]
    [SetCulture("tr-TR")]
    public void HighlightText_uses_ordinal_case_matching(bool ignoreCase, int expected, string selected)
    {
        // Arrange
        using var pres = ReplacementPresentation(new A.Paragraph(new A.Run(new A.Text("FILE file"))));
        var box = pres.Slides[0].Shapes.Last().TextBox;

        // Act
        var count = box.HighlightText("file", new Color("FFFF00"), ignoreCase);

        // Assert
        count.Should().Be(expected);
        string.Concat(box.Paragraphs[0].Portions.Where(portion => !portion.TextHighlightColor.IsTransparent).Select(portion => portion.Text))
            .Should().Be(selected);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void HighlightText_does_not_cross_paragraphs_breaks_or_fields(bool regex)
    {
        // Arrange
        using var pres = ReplacementPresentation(
            new A.Paragraph(new A.Run(new A.Text("ab"))), new A.Paragraph(new A.Run(new A.Text("cd"))),
            new A.Paragraph(new A.Run(new A.Text("ab")), new A.Break(), new A.Run(new A.Text("cd"))),
            new A.Paragraph(new A.Run(new A.Text("ab")),
                new A.Field(new A.Text("FIELD")) { Id = "{E7559270-9838-4E40-AB23-A20A96371411}", Type = "slidenum" },
                new A.Run(new A.Text("cd"))));
        var box = pres.Slides[0].Shapes.Last().TextBox;
        var before = box.Text;

        // Act
        var count = regex ? box.HighlightRegex(new Regex("abcd|FIELD"), Color.Black)
            : box.HighlightText("abcd", Color.Black) + box.HighlightText("FIELD", Color.Black);

        // Assert
        count.Should().Be(0);
        box.Text.Should().Be(before);
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var sdk = PresentationDocument.Open(saved, false);
        sdk.PresentationPart!.SlideParts.Single().Slide.Descendants<A.Highlight>().Should().BeEmpty();
        ValidatePresentation(pres);
    }

    [TestCase(RegexOptions.None)]
    [TestCase(RegexOptions.RightToLeft)]
    public void HighlightRegex_supports_options_cross_run_matches_and_ignores_empty_matches(RegexOptions options)
    {
        // Arrange
        using var pres = ReplacementPresentation(new A.Paragraph(
            new A.Run(new A.Text("x1")), new A.Run(new A.Text("2 X3 tail"))));
        var box = pres.Slides[0].Shapes.Last().TextBox;
        var regex = new Regex(@"x\d+|(?=tail)", options | RegexOptions.IgnoreCase, TimeSpan.FromSeconds(1));

        // Act
        var count = box.HighlightRegex(regex, new Color("FFFF00"));

        // Assert
        count.Should().Be(2);
        box.Text.Should().Be("x12 X3 tail");
        string.Concat(box.Paragraphs[0].Portions.Where(portion => !portion.TextHighlightColor.IsTransparent).Select(portion => portion.Text))
            .Should().Be("x12X3");
        ValidatePresentation(pres);
    }

    [Test]
    public void HighlightRegex_merges_adjacent_matches_without_breaking_Unicode_surrogate_pairs()
    {
        // Arrange
        using var pres = ReplacementPresentation(new A.Paragraph(new A.Run(new A.Text("a😀b"))));
        var box = pres.Slides[0].Shapes.Last().TextBox;

        // Act
        var count = box.HighlightRegex(new Regex("."), new Color("FFFF00"));

        // Assert
        count.Should().Be(4);
        box.Text.Should().Be("a😀b");
        box.Paragraphs[0].Portions.Should().ContainSingle();
        ValidatePresentation(pres);
    }

    [Test]
    public void HighlightRegex_rejects_surrogate_splitting_without_mutating_the_text_box()
    {
        // Arrange
        using var pres = ReplacementPresentation(new A.Paragraph(new A.Run(new A.Text("abc"))),
            new A.Paragraph(new A.Run(new A.Text("😀"))));
        var box = pres.Slides[0].Shapes.Last().TextBox;
        var before = box.Text;

        // Act
        var highlight = () => box.HighlightRegex(new Regex("abc|\\uD83D"), Color.Black);

        // Assert
        highlight.Should().Throw<ArgumentException>();
        box.Text.Should().Be(before);
        box.Paragraphs[0].Portions[0].TextHighlightColor.IsTransparent.Should().BeTrue();
        ValidatePresentation(pres);
    }

    [Test]
    public void HighlightRegex_timeout_does_not_partially_update_a_text_box()
    {
        // Arrange
        using var pres = ReplacementPresentation(new A.Paragraph(new A.Run(new A.Text("abc"))),
            new A.Paragraph(new A.Run(new A.Text(new string('a', 10_000) + "!"))));
        var box = pres.Slides[0].Shapes.Last().TextBox;
        var regex = new Regex("abc|(a+)+$", RegexOptions.None, TimeSpan.FromMilliseconds(1));

        // Act
        var highlight = () => box.HighlightRegex(regex, Color.Black);

        // Assert
        highlight.Should().Throw<RegexMatchTimeoutException>();
        box.Paragraphs[0].Portions[0].TextHighlightColor.IsTransparent.Should().BeTrue();
        box.Paragraphs[0].Text.Should().Be("abc");
    }

    [Test]
    public void HighlightText_preserves_outside_highlights_and_can_clear_selected_direct_highlights()
    {
        // Arrange
        using var pres = ReplacementPresentation(new A.Paragraph(new A.Run(new A.Text("abcde"))));
        var box = pres.Slides[0].Shapes.Last().TextBox;
        box.Paragraphs[0].Portions[0].TextHighlightColor = new Color("FF0000");

        // Act
        var count = box.HighlightText("bcd", Color.NoColor);

        // Assert
        count.Should().Be(1);
        var portions = box.Paragraphs[0].Portions.ToArray();
        portions.Select(portion => portion.Text).Should().Equal("a", "bcd", "e");
        portions[0].TextHighlightColor.Hex.Should().Be("FF0000");
        portions[1].TextHighlightColor.IsTransparent.Should().BeTrue();
        portions[2].TextHighlightColor.Hex.Should().Be("FF0000");
        ValidatePresentation(pres);
    }

    [Test]
    public void HighlightText_round_trips_alpha_and_repeated_calls_do_not_duplicate_highlights()
    {
        // Arrange
        using var pres = ReplacementPresentation(new A.Paragraph(new A.Run(new A.Text("abc"))));
        var box = pres.Slides[0].Shapes.Last().TextBox;
        var color = new Color("FFFF00") { Alpha = 128 };
        box.HighlightText("b", color);

        // Act
        box.HighlightText("b", color);

        // Assert
        using var saved = new MemoryStream();
        pres.Save(saved);
        using (var sdk = PresentationDocument.Open(saved, false))
        {
            var highlight = sdk.PresentationPart!.SlideParts.Single().Slide.Descendants<A.Highlight>().Single();
            highlight.RgbColorModelHex!.GetFirstChild<A.Alpha>()!.Val!.Value.Should().Be(50196);
        }

        using var reopened = new Presentation(saved);
        var portions = reopened.Slides[0].Shapes.Last().TextBox.Paragraphs[0].Portions;
        portions.Should().HaveCount(3);
        portions[1].TextHighlightColor.Alpha.Should().BeApproximately(128, 0.001f);
        ValidatePresentation(reopened);
    }

    [TestCase(null)]
    [TestCase("")]
    public void HighlightText_rejects_invalid_queries(string? query)
    {
        // Arrange
        using var pres = ReplacementPresentation(new A.Paragraph(new A.Run(new A.Text("abc"))));
        var box = pres.Slides[0].Shapes.Last().TextBox;

        // Act
        var highlight = () => box.HighlightText(query!, Color.Black);

        // Assert
        highlight.Should().Throw<ArgumentException>();
        box.Paragraphs[0].Portions.Should().ContainSingle();
        box.Text.Should().Be("abc");
    }

    [TestCase(float.NaN)]
    [TestCase(float.PositiveInfinity)]
    [TestCase(-1)]
    [TestCase(256)]
    public void HighlightText_rejects_invalid_alpha_without_mutation(float alpha)
    {
        // Arrange
        using var pres = ReplacementPresentation(new A.Paragraph(new A.Run(new A.Text("abc"))));
        var box = pres.Slides[0].Shapes.Last().TextBox;

        // Act
        var highlight = () => box.HighlightText("b", new Color("FFFF00") { Alpha = alpha });

        // Assert
        highlight.Should().Throw<ArgumentOutOfRangeException>();
        box.Paragraphs[0].Portions.Should().ContainSingle();
    }

    [Test]
    public void ReplaceText_matches_across_runs_and_preserves_surrounding_formatting()
    {
        // Arrange
        using var pres = ReplacementPresentation(new A.Paragraph(
            new A.ParagraphProperties { Level = 1 },
            new A.Run(new A.RunProperties { Bold = true, FontSize = 1800 }, new A.Text("pre XY")),
            new A.Run(new A.RunProperties { Italic = true, FontSize = 2400 }, new A.Text("Z tail"))));
        var textBox = pres.Slides[0].Shapes.Last().TextBox!;

        // Act
        var count = textBox.ReplaceText("XYZ", "Q");

        // Assert
        count.Should().Be(1);
        textBox.Text.Should().Be("pre Q tail");
        using var saved = new MemoryStream();
        pres.Save(saved);
        using (var document = PresentationDocument.Open(saved, false))
        {
            var paragraph = document.PresentationPart!.SlideParts.Single().Slide.Descendants<P.Shape>().Last().TextBody!.Elements<A.Paragraph>().Single();
            paragraph.ParagraphProperties!.Level!.Value.Should().Be(1);
            var runs = paragraph.Elements<A.Run>().ToArray();
            runs[0].Text!.Text.Should().Be("pre Q");
            runs[0].RunProperties!.Bold!.Value.Should().BeTrue();
            runs[0].RunProperties.FontSize!.Value.Should().Be(1800);
            runs[1].Text!.Text.Should().Be(" tail");
            runs[1].RunProperties!.Italic!.Value.Should().BeTrue();
            runs[1].RunProperties.FontSize!.Value.Should().Be(2400);
        }

        using var reopened = new Presentation(saved);
        reopened.Slides[0].Shapes.Last().TextBox!.Text.Should().Be("pre Q tail");
        ValidatePresentation(reopened);
    }

    [TestCase("aaaa", "aa", "a", 2, "aa")]
    [TestCase("aaaa", "aa", "aaa", 2, "aaaaaa")]
    [TestCase("ababa ba", "ba", "X", 3, "aXX X")]
    [TestCase("aaa", "a", "", 3, "")]
    [TestCase("abc", "missing", "X", 0, "abc")]
    [TestCase("aXYz", "XY", "XY!", 1, "aXY!z")]
    [TestCase("aXYz", "XY", "  <&😀>  ", 1, "a  <&😀>  z")]
    public void ReplaceText_replaces_original_non_overlapping_matches(string original, string oldText, string newText, int expectedCount, string expected)
    {
        // Arrange
        var middle = original.Length / 2;
        using var pres = ReplacementPresentation(new A.Paragraph(
            new A.Run(new A.Text(original[..middle])), new A.Run(new A.Text(string.Empty)),
            new A.Run(new A.Text(original[middle..]))));
        var textBox = pres.Slides[0].Shapes.Last().TextBox!;

        // Act
        var count = textBox.ReplaceText(oldText, newText);

        // Assert
        count.Should().Be(expectedCount);
        textBox.Text.Should().Be(expected);
        ValidatePresentation(pres);
    }

    [TestCase(false, 1, "FILE X")]
    [TestCase(true, 2, "X X")]
    [SetCulture("tr-TR")]
    public void ReplaceText_uses_ordinal_case_matching(bool ignoreCase, int expectedCount, string expected)
    {
        // Arrange
        using var pres = ReplacementPresentation(new A.Paragraph(new A.Run(new A.Text("FILE file"))));
        var textBox = pres.Slides[0].Shapes.Last().TextBox!;

        // Act
        var count = textBox.ReplaceText("file", "X", ignoreCase);

        // Assert
        count.Should().Be(expectedCount);
        textBox.Text.Should().Be(expected);
    }

    [Test]
    public void ReplaceText_inserts_soft_breaks_for_multiline_values()
    {
        // Arrange
        using var pres = ReplacementPresentation(new A.Paragraph(
            new A.Run(new A.RunProperties { Bold = true }, new A.Text("aXYz"))));
        var textBox = pres.Slides[0].Shapes.Last().TextBox!;

        // Act
        var count = textBox.ReplaceText("XY", "Q\r\nR\nS\rT");

        // Assert
        count.Should().Be(1);
        textBox.Text.Should().Be(string.Join(Environment.NewLine, "aQ", "R", "S", "Tz"));
        using var saved = new MemoryStream();
        pres.Save(saved);
        using (var document = PresentationDocument.Open(saved, false))
        {
            var body = document.PresentationPart!.SlideParts.Single().Slide.Descendants<P.Shape>().Last().TextBody!;
            body.Elements<A.Paragraph>().Should().HaveCount(1);
            body.Descendants<A.Break>().Should().HaveCount(3);
            body.Descendants<A.Run>().Should().OnlyContain(run => run.RunProperties!.Bold!.Value);
        }

        ValidatePresentation(pres);
    }

    [Test]
    public void ReplaceText_does_not_cross_paragraphs_breaks_or_fields()
    {
        // Arrange
        using var pres = ReplacementPresentation(
            new A.Paragraph(new A.Run(new A.Text("ab"))),
            new A.Paragraph(new A.Run(new A.Text("cd"))),
            new A.Paragraph(new A.Run(new A.Text("ab")), new A.Break(), new A.Run(new A.Text("cd"))),
            new A.Paragraph(new A.Run(new A.Text("ab")),
                new A.Field(new A.Text("FIELD")) { Id = "{E7559270-9838-4E40-AB23-A20A96371411}", Type = "slidenum" },
                new A.Run(new A.Text("cd"))));
        var textBox = pres.Slides[0].Shapes.Last().TextBox!;

        // Act
        var count = textBox.ReplaceText("abcd", "X") + textBox.ReplaceText("FIELD", "X");

        // Assert
        count.Should().Be(0);
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var document = PresentationDocument.Open(saved, false);
        document.PresentationPart!.SlideParts.Single().Slide.Descendants<A.Field>().Single().InnerText.Should().Be("FIELD");
        ValidatePresentation(pres);
    }

    [TestCase(null, "X")]
    [TestCase("", "X")]
    [TestCase("a", null)]
    public void ReplaceText_rejects_invalid_arguments_without_mutation(string? oldText, string? newText)
    {
        // Arrange
        using var pres = ReplacementPresentation(new A.Paragraph(new A.Run(new A.Text("abc"))));
        var textBox = pres.Slides[0].Shapes.Last().TextBox!;

        // Act
        var replace = () => textBox.ReplaceText(oldText!, newText!);

        // Assert
        replace.Should().Throw<ArgumentException>();
        textBox.Text.Should().Be("abc");
    }

    [Test]
    public void ReplaceText_rejects_invalid_XML_characters_before_mutation()
    {
        // Arrange
        using var pres = ReplacementPresentation(new A.Paragraph(new A.Run(new A.Text("abc"))));
        var textBox = pres.Slides[0].Shapes.Last().TextBox!;

        // Act
        var replace = () => textBox.ReplaceText("a", "\0");

        // Assert
        replace.Should().Throw<System.Xml.XmlException>();
        textBox.Text.Should().Be("abc");
    }

    private static Presentation ReplacementPresentation(params A.Paragraph[] paragraphs)
    {
        using var source = new Presentation(p => p.Slide());
        source.Slides[0].Shapes.AddTextBox(0, 0, 300, 100, string.Empty);
        using var saved = new MemoryStream();
        source.Save(saved);
        using (var document = PresentationDocument.Open(saved, true))
        {
            var body = document.PresentationPart!.SlideParts.Single().Slide.Descendants<P.Shape>().Last().TextBody!;
            body.RemoveAllChildren<A.Paragraph>();
            body.Append(paragraphs);
        }

        return new Presentation(saved);
    }

    [Test]
    public void Text_Getter_returns_text_of_table_Cell()
    {
        // Arrange
        var textBox1 = new Presentation(TestAsset("008.pptx")).Slide(1).Shape(3).TextBox;
        var textBox2 = new Presentation(TestAsset("001.pptx")).Slide(2).Shape(3).Table.Rows[0].Cells[0]
            .TextBox;
        var textBox3 = new Presentation(TestAsset("009_table.pptx")).Slide(3).Shape(3).Table.Rows[0].Cells[0].TextBox;

        // Act
        var text1 = textBox1.Text;
        var text2 = textBox2.Text;
        var text3 = textBox3.Text;

        // Assert
        text1.Should().NotBeEmpty();
        text2.Should().BeEquivalentTo("id3");
        text3.Should().BeEquivalentTo($"0:0_p1_lvl1{Environment.NewLine}0:0_p2_lvl2");
    }

    [Test]
    public void Text_Getter_returns_text_from_New_Slide()
    {
        // Arrange
        var pres = new Presentation(TestAsset("031.pptx"));
        var layout = pres.SlideMaster(1).SlideLayout(1);

        // Act
        pres.Slides.Add(layout.Number);
        var textBox = pres.Slides.Last().Shape("Holder 2").TextBox;
        var text = textBox!.Text;

        // Assert
        text.Should().BeEquivalentTo("");
    }

    // Regression: https://github.com/ShapeCrawler/ShapeCrawler/issues/1284
    [Test]
    public void SetText_on_title_placeholder_after_reseeding_slides_from_layout_does_not_throw()
    {
        // Arrange
        var pres = new Presentation(TestAsset("086.pptx"));
        for (var i = pres.Slides.Count; i >= 1; i--) pres.Slide(i).Remove();

        var layout = pres.MasterSlides[0].SlideLayout("Title and Content");
        pres.Slides.Add(layout.Number);
        var slide = pres.Slide(pres.Slides.Count);
        var titleShape = slide.Shapes.First(s => s.PlaceholderType == PlaceholderType.Title);

        // Act
        var act = () => titleShape.TextBox!.SetText("Hello");

        // Assert
        act.Should().NotThrow();
        ValidatePresentation(pres);
    }

    [Test]
    public void SetText_can_update_content_multiple_times()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-case005_text-frame.pptx"));
        var textBox = pres.Slides[0].Shape("TextBox 1").TextBox!;
        var modifiedPres = new MemoryStream();

        // Act
        var newText = textBox.Text.Replace("{{replace_this}}", "confirm this");
        textBox.SetText(newText);
        newText = textBox.Text.Replace("{{replace_that}}", "confirm that");
        textBox.SetText(newText);

        // Assert
        pres.Save(modifiedPres);
        pres = new Presentation(modifiedPres);
        textBox = pres.Slides[0].Shapes.Shape<IShape>("TextBox 1").TextBox;
        textBox.Text.Should().Contain("confirm this");
        textBox.Text.Should().Contain("confirm that");
    }

    [Test]
    public void SetText_reduces_font_size_When_text_is_Overflow()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var textBox = pres.Slide(1).Shape("TextBox 8").TextBox!;

        // Act
        textBox.SetText("Shrink text on overflow");

        // Assert
        textBox.Text.Should().BeEquivalentTo("Shrink text on overflow");
        textBox.Paragraphs[0].Portions[0].Font!.Size.Should().BeApproximately(8, 1);
    }

    [Test]
    public void SetText_preserves_font()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var textBox = pres.Slide(1).Shape("TextBox 8").TextBox!;
        var expectedFont = textBox.Paragraphs[0].Portions[0].Font!.LatinName;

        // Act
        textBox.SetText("Shrink text on overflow");

        // Assert
        textBox.Paragraphs[0].Portions[0].Font!.LatinName.Should().BeEquivalentTo(expectedFont);
    }

    [Test]
    public void SetText_preserves_font_color()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-case003.pptx"));
        var shape = pres.Slide(1).Shape("AutoShape 4");
        var textBox = shape.TextBox!;
        var expectedColor = textBox.Paragraphs[0].Portions[0].Font!.Color.Hex;

        // Act
        textBox.SetText("AutoShape 4 some text");

        // Assert
        textBox.Paragraphs[0].Portions[0].Font!.Color.Hex.Should().Be(expectedColor);
    }

    [TestCase(1, "TextBox 3")]
    [TestCase(3, "Text Placeholder 3")]
    public void SetText_preserves_original_font_size(int slideNumber, string shapeName)
    {
        // Arrange
        using var pres = new Presentation(TestAsset("001.pptx"));
        var textBox = pres.Slide(slideNumber).Shape(shapeName).TextBox!;
        var expectedSize = textBox.Paragraphs[0].Portions[0].Font!.Size;

        // Act
        textBox.SetText("New text");

        // Assert
        textBox.Paragraphs[0].Portions[0].Font!.Size.Should().Be(expectedSize);
        using var reopened = SaveAndOpenPresentation(pres);
        reopened.Slide(slideNumber).Shape(shapeName).TextBox!.Paragraphs[0].Portions[0].Font!.Size
            .Should().Be(expectedSize);
    }

    [TestCase(AutofitType.None, 15.5)]
    [TestCase(AutofitType.Resize, 15.5)]
    [TestCase(AutofitType.Shrink, 15)]
    public void SetText_preserves_font_size_for_each_replacement_paragraph(AutofitType autofitType, double size)
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slide(1).Shapes;
        shapes.AddShape(100, 100, 600, 300);
        var textBox = shapes.Last().TextBox!;
        textBox.SetText("Original text");
        textBox.AutofitType = autofitType;
        textBox.Paragraphs[0].Portions[0].Font!.Size = (decimal)size;
        textBox.Paragraphs[0].Portions.AddText(" differently formatted text");
        textBox.Paragraphs[0].Portions[1].Font!.Size = 30;
        var replacement = $"New text{Environment.NewLine}Second line";

        // Act
        textBox.SetText(replacement);

        // Assert
        textBox.Text.Should().Be(replacement);
        textBox.Paragraphs.Should().HaveCount(2);
        foreach (var paragraph in textBox.Paragraphs)
        {
            paragraph.Portions.Should().ContainSingle().Which.Font!.Size.Should().Be((decimal)size);
        }

        using var reopened = SaveAndOpenPresentation(pres);
        var reopenedTextBox = reopened.Slide(1).Shapes.Last().TextBox!;
        reopenedTextBox.Text.Should().Be(replacement);
        reopenedTextBox.Paragraphs.SelectMany(p => p.Portions).Should()
            .OnlyContain(p => p.Font!.Size == (decimal)size);
        ValidatePresentation(pres);
    }

    [Test]
    public void SetMarkdownText()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slide(1).Shapes;
        shapes.AddShape(100, 100, 200, 200);
        var shape = shapes.Last();
        var textBox = shape.TextBox!;

        // Act
        textBox.SetMarkdownText("**Hello** World!");

        // Assert
        textBox.Paragraphs[0].Portions[0].Font!.IsBold.Should().BeTrue();
        textBox.Paragraphs[0].Portions[1].Font!.IsBold.Should().BeFalse();
    }

    [Test]
    public void SetMarkdownText_sets_list()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slide(1).Shapes;
        shapes.AddShape(100, 100, 200, 200);
        var shape = shapes.Last();
        var textBox = shape.TextBox!;
        var list = """
                   - Item 1
                   - Item 2
                   """;
        // Act
        textBox.SetMarkdownText(list);

        // Assert
        textBox.Paragraphs[0].Bullet.Type.Should().NotBe(BulletType.None);
        ValidatePresentation(pres);
    }

    [Test]
    [Platform(Exclude = "Linux,MacOSX", Reason = "Test fails")]
    public void SetText_resizes_shape_to_fit_text()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-case003.pptx"));
        var shape = pres.Slide(1).Shape("AutoShape 4");
        var textBox = shape.TextBox!;
        using var expectedPres = new Presentation(TestAsset("autoshape-case003.pptx"));
        var expectedShape = expectedPres.Slide(1).Shape("AutoShape 4");
        expectedShape.TextBox!.Paragraphs[0].Text = "AutoShape 4 some text";
        var expectedFontSize = textBox.Paragraphs[0].Portions[0].Font!.Size;

        // Act
        textBox.SetText("AutoShape 4 some text");

        // Assert
        textBox.Text.Should().Be("AutoShape 4 some text");
        textBox.Paragraphs[0].Portions[0].Font!.Size.Should().Be(expectedFontSize);
        shape.Height.Should().BeApproximately(expectedShape.Height, 1m);
        shape.Y.Should().BeApproximately(expectedShape.Y, 2m);
        ValidatePresentation(pres);
    }

    [Test]
    [Platform(Exclude = "Linux,MacOSX", Reason = "Test fails on Ubuntu and macOS")]
    public void SetText_updates_shape_height_to_fit_text()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-case003.pptx"));
        var shape = pres.Slide(1).Shape("AutoShape 4");
        var textBox = shape.TextBox!;
        var originalHeight = shape.Height;

        // Act
        textBox.Paragraphs.Add();
        textBox.Paragraphs.Last().Text = "AutoShape 4 some text";
        textBox.Paragraphs.Add();
        textBox.Paragraphs.Last().Text = "AutoShape 4 some text";
        textBox.Paragraphs.Add();
        textBox.Paragraphs.Last().Text = "AutoShape 4 some text";

        // Assert
        shape.Height.Should().BeGreaterThan(originalHeight);
        ValidatePresentation(pres);
    }

    [Test]
    public void Paragraph_Text_Setter_updates_shape_height_to_fit_text_with_line_breaks()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-case003.pptx"));
        var shape = pres.Slide(1).Shape("AutoShape 4");
        var paragraph = shape.TextBox!.Paragraphs[0];
        paragraph.Text = "One";
        var oneLineHeight = shape.Height;

        // Act
        paragraph.Text = $"One{Environment.NewLine}Two{Environment.NewLine}Three";

        // Assert
        shape.Height.Should().BeGreaterThan(oneLineHeight);
        ValidatePresentation(pres);
    }

    [Test]
    public void SetText_sets_text_for_New_Shape()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slides[0].Shapes;
        shapes.AddShape(50, 60, 100, 70);
        var textBox = shapes.Last().TextBox!;

        // Act
        textBox.SetText("Test");

        // Assert
        textBox.Text.Should().Be("Test");
        ValidatePresentation(pres);
    }

    [Test]
    [Platform(Exclude = "Linux", Reason = "Test fails on ubuntu-latest")]
    public void AutofitType_Setter_resizes_width()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-case003.pptx"));
        var shape = pres.Slide(1).Shape("AutoShape 6");
        var textBox = shape.TextBox!;

        // Act
        textBox.AutofitType = AutofitType.Resize;

        // Assert
        shape.Width.Should().BeApproximately(103m, 103m);
        ValidatePresentation(pres);
    }

    [Test]
    public void AutofitType_Setter_updates_height()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-case003.pptx"));
        var shape = pres.Slide(1).Shape("AutoShape 7");
        var textBox = shape.TextBox!;

        // Act
        textBox.AutofitType = AutofitType.Resize;

        // Assert
        shape.Height.Should().BeApproximately(30m, 30m);
        ValidatePresentation(pres);
    }

    [Test]
    public void AutofitType_Getter_returns_text_autofit_type()
    {
        // Arrange
        var pptx = TestAsset("001.pptx");
        var pres = new Presentation(pptx);
        var autoShape = pres.Slides[0].Shapes.GetById<IShape>(9);
        var textBox = autoShape.TextBox;

        // Act
        var autofitType = textBox.AutofitType;

        // Assert
        autofitType.Should().Be(AutofitType.Shrink);
    }

    [Test]
    public void Shape_IsAutoShape()
    {
        // Arrange
        var shapeCase1 = new Presentation(TestAsset("008.pptx")).Slides[0].Shapes.First(sp => sp.Id == 3);
        var shapeCase2 = new Presentation(TestAsset("021.pptx")).Slides[3].Shapes.First(sp => sp.Id == 2);
        var shapeCase3 = new Presentation(TestAsset("011_dt.pptx")).Slides[0].Shapes.First(sp => sp.Id == 54275);

        // Act
        var autoshapecase1 = shapeCase1;
        var autoshapecase2 = shapeCase2;
        var autoshapecase3 = shapeCase3;

        // Assert
        autoshapecase1.Should().NotBeNull();
        autoshapecase2.Should().NotBeNull();
        autoshapecase3.Should().NotBeNull();
    }

    [Test]
    public void Paragraphs_Add_adds_new_text_paragraph_at_the_end_And_returns_added_paragraph()
    {
        // Arrange
        const string TEST_TEXT = "ParagraphsAdd";
        var mStream = new MemoryStream();
        var pres = new Presentation(TestAsset("001.pptx"));
        var textFrame = pres.Slides[0].Shapes.First(sp => sp.Id == 4).TextBox;
        var originParagraphsCount = textFrame.Paragraphs.Count;

        // Act
        textFrame.Paragraphs.Add();
        var addedPara = textFrame.Paragraphs.Last();
        addedPara.Text = TEST_TEXT;

        // Assert
        var lastPara = textFrame.Paragraphs.Last();
        lastPara.Text.Should().BeEquivalentTo(TEST_TEXT);
        textFrame.Paragraphs.Should().HaveCountGreaterThan(originParagraphsCount);

        pres.Save(mStream);
        pres = new Presentation(mStream);
        textFrame = pres.Slides[0].Shapes.First(sp => sp.Id == 4).TextBox;
        textFrame.Paragraphs.Last().Text.Should().BeEquivalentTo(TEST_TEXT);
        textFrame.Paragraphs.Should().HaveCountGreaterThan(originParagraphsCount);
    }

    [Test]
    public void Paragraphs_Add_adds_paragraph()
    {
        // Arrange
        var pptxStream = TestAsset("autoshape-case007.pptx");
        var pres = new Presentation(pptxStream);
        var paragraphs = pres.Slides[0].Shapes.Shape<IShape>("AutoShape 1").TextBox.Paragraphs;

        // Act
        paragraphs.Add();

        // Assert
        paragraphs.Should().HaveCount(6);
    }

    [Test]
    public void
        Paragraphs_Add_adds_new_text_paragraph_at_the_end_And_returns_added_paragraph_When_it_has_been_added_after_text_frame_changed()
    {
        var pres = new Presentation(TestAsset("001.pptx"));
        var autoShape = pres.Slides[0].Shapes.First(sp => sp.Id == 3);
        var textBox = autoShape.TextBox!;
        var paragraphs = textBox.Paragraphs;
        // Act
        textBox.SetText("A new text");
        paragraphs.Add();
        var addedParagraph = paragraphs.Last();

        // Assert
        addedParagraph.Should().NotBeNull();
    }

    [Test]
    [TestCase("autoshape-case003.pptx", 1, "AutoShape 7")]
    [TestCase("001.pptx", 1, "Head 1")]
    [TestCase("autoshape-case014.pptx", 1, "Content Placeholder 1")]
    public void AutofitType_Setter_sets_autofit_type(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(file));
        var shape = pres.Slides[slideNumber - 1].Shapes.Shape(shapeName);
        var autoShape = shape;
        var textFrame = autoShape.TextBox!;

        // Act
        textFrame.AutofitType = AutofitType.Resize;

        // Assert
        textFrame.AutofitType.Should().Be(AutofitType.Resize);
        ValidatePresentation(pres);
    }

    [Test]
    public void SetText_sets_long_text()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-case013.pptx"));
        var shape = pres.Slide(1).Shape("AutoShape 1");
        var textBox = shape.TextBox!;
        using var expectedPres = new Presentation(TestAsset("autoshape-case013.pptx"));
        var expectedShape = expectedPres.Slide(1).Shape("AutoShape 1");
        expectedShape.TextBox!.Paragraphs[0].Text = "Some sentence. Some sentence";

        // Act
        textBox.SetText("Some sentence. Some sentence");

        // Assert
        textBox.Text.Should().Be("Some sentence. Some sentence");
        shape.Height.Should().BeApproximately(expectedShape.Height, 1m);
    }

    [Test]
    [SlideShape("009_table.pptx", 4, 2, "Title text")]
    [SlideShape("001.pptx", 1, 5, " id5-Text1")]
    [SlideShape("019.pptx", 1, 2, "1")]
    [SlideShape("014.pptx", 2, 5, "Test subtitle")]
    [SlideShape("011_dt.pptx", 1, 54275, "Jan 2018")]
    [SlideShape("021.pptx", 4, 2, "test footer")]
    [SlideShape("012_title-placeholder.pptx", 1, 2, "Test title text")]
    [SlideShape("012_title-placeholder.pptx", 1, 3, "P1 P2")]
    public void Text_Getter_returns_text(IShape shape, string expectedText)
    {
        // Arrange
        var textFrame = shape.TextBox;

        // Act
        var text = textFrame.Text;

        // Assert
        text.Should().BeEquivalentTo(expectedText);
    }

    [Test]
    [SlideShape("014.pptx", 2, 5, TextVerticalAlignment.Middle)]
    public void VerticalAlignment_Getter_returns_vertical_alignment(IShape shape,
        TextVerticalAlignment expectedVAlignment)
    {
        // Arrange
        var textBox = shape.TextBox;

        // Act-Assert
        textBox.VerticalAlignment.Should().Be(expectedVAlignment);
    }

    [Test]
    [SlideShape("001.pptx", 1, 6, "id6-Text1#NewLine#Text2")]
    [SlideShape("014.pptx", 1, 61, "test1#NewLine#test2#NewLine#test3#NewLine#test4#NewLine#test5")]
    [SlideShape("011_dt.pptx", 1, 2, "P1#NewLine#")]
    public void Text_Getter_returns_text_with_New_Line(IShape shape, string expectedText)
    {
        // Arrange
        expectedText = expectedText.Replace("#NewLine#", Environment.NewLine);
        var textFrame = shape.TextBox;

        // Act
        var text = textFrame.Text;

        // Assert
        text.Should().BeEquivalentTo(expectedText);
    }

    [Test]
    public void Text_Getter_returns_text_box_text()
    {
        // Arrange
        var pres = new Presentation(pres => { pres.Slide(slide => { slide.TextShape("Hello World!"); }); });
        var textBox = pres.Slide(1).Shapes.First().TextBox!;

        // Act-Assert
        textBox.Text.Should().Be("Hello World!");
    }

    [Test]
    [TestCase("001.pptx", 1, "TextBox 2")]
    [TestCase("020.pptx", 3, "TextBox 7")]
    [TestCase("001.pptx", 2, "Header 1")]
    [TestCase("autoshape-case004_subtitle.pptx", 1, "Subtitle 1")]
    [TestCase("autoshape-case008_text-frame.pptx", 1, "AutoShape 1")]
    public void SetText_updates_content(string presName, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(presName));
        var textBox = pres.Slides[slideNumber - 1].Shapes.Shape<IShape>(shapeName).TextBox!;
        var mStream = new MemoryStream();

        // Act
        textBox.SetText("Test");

        // Assert
        textBox.Text.Should().BeEquivalentTo("Test");
        textBox.Paragraphs.Should().HaveCount(1);

        pres.Save(mStream);
        pres = new Presentation(mStream);
        textBox = pres.Slides[slideNumber - 1].Shapes.Shape<IShape>(shapeName).TextBox;
        textBox.Text.Should().BeEquivalentTo("Test");
        textBox.Paragraphs.Should().HaveCount(1);
    }

    [Test]
    [SlideShape("autoshape-case012.pptx", 1, "Shape 1")]
    public void SetText(IShape shape)
    {
        // Arrange
        var textBox = shape.TextBox!;

        // Act
        textBox.SetText("some text");

        // Assert
        textBox.Text.Should().BeEquivalentTo("some text");
    }

    [Test]
    public void SetText_adds_two_paragraphs_in_the_table_cell()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slide(1).Shapes;
        shapes.AddTable(50, 50, 2, 2);
        var cellTextBox = pres.Slide(1).Shapes.First().Table[0, 0].TextBox;

        // Act
        cellTextBox.SetText($"Text 1{Environment.NewLine}Text 2");

        // Assert
        cellTextBox.Paragraphs[0].Text.Should().BeEquivalentTo("Text 1");
        cellTextBox.Paragraphs[1].Text.Should().BeEquivalentTo("Text 2");
    }

    [Test]
    [SlideShape("073 replacing text.pptx", 1, "TextBox 3")]
    public void SetText_preserves_new_lines(IShape shape)
    {
        // Arrange
        var expectedText = "Hello" + Environment.NewLine + Environment.NewLine + "Earth";
        var textBox = shape.TextBox!;

        // Act
        var newText = textBox.Text.Replace("World", "Earth");
        textBox.SetText(newText);

        // Assert
        textBox.Text.Should().Be(expectedText);
    }

    [Test]
    [SlideShape("autoshape-case003.pptx", 1, "AutoShape 6", false)]
    [SlideShape("autoshape-case003.pptx", 1, "AutoShape 2", true)]
    [SlideShape("autoshape-case013.pptx", 1, "AutoShape 1", true)]
    public void TextWrapped_Getter_returns_value_indicating_whether_text_is_wrapped_in_shape(IShape shape,
        bool isTextWrapped)
    {
        // Arrange
        var autoShape = shape;
        var textFrame = autoShape.TextBox!;

        // Act
        var textWrapped = textFrame.TextWrapped;

        // Assert
        textWrapped.Should().Be(isTextWrapped);
    }

    [Test]
    [SlideShape("009_table.pptx", 3, 2, 1)]
    [SlideShape("020.pptx", 3, 8, 2)]
    [SlideShape("001.pptx", 2, 2, 1)]
    public void Paragraphs_Count_returns_number_of_paragraphs_in_the_text_box(IShape shape,
        int expectedParagraphsCount)
    {
        // Arrange
        var textFrame = shape.TextBox;

        // Act
        var paragraphsCount = textFrame.Paragraphs.Count;

        // Assert
        paragraphsCount.Should().Be(expectedParagraphsCount);
    }

    [Test]
    public void Paragraphs_Count_returns_number_of_paragraphs_in_the_table_cell_text_box()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));
        var textFrame = pres.Slide(3).Shape(3).Table.Rows[0].Cells[0].TextBox;

        // Act
        var paragraphsCount = textFrame.Paragraphs.Count;

        // Assert
        paragraphsCount.Should().Be(2);
    }

    [Test]
    [SlideShape("autoshape-case003.pptx", 1, "AutoShape 2", 7.09)]
    [SlideShape("autoshape-case003.pptx", 1, "AutoShape 3", 8.50)]
    public void LeftMargin_getter_returns_left_margin(IShape shape, double expectedMargin)
    {
        // Arrange
        var expectedMarginDecimal = (decimal)expectedMargin;

        // Act & Assert
        shape.TextBox.LeftMargin.Should().BeApproximately(expectedMarginDecimal, 0.01m);
    }

    [Test]
    [SlideShape("autoshape-case003.pptx", 1, "AutoShape 2")]
    public void LeftMargin_setter_sets_left_margin_of_text_frame_in_centimeters(IShape shape)
    {
        // Arrange
        var textBox = shape.TextBox;

        // Act
        textBox.LeftMargin = 10m;

        // Assert
        textBox.LeftMargin.Should().Be(10m);
    }

    [Test]
    [SlideShape("autoshape-case003.pptx", 1, "AutoShape 2", 7.09)]
    public void RightMargin_getter_returns_right_margin(IShape shape, double expectedMargin)
    {
        // Arrange
        var expectedMarginDecimal = (decimal)expectedMargin;

        // Act & Assert
        shape.TextBox.RightMargin.Should().Be(expectedMarginDecimal);
    }

    [Test]
    [SlideShape("autoshape-case003.pptx", 1, "AutoShape 2", 3.69)]
    [SlideShape("autoshape-case003.pptx", 1, "AutoShape 3", 3.96)]
    public void TopMargin_getter_returns_top_margin_of_text_frame_in_centimeters(IShape shape,
        double expectedMargin)
    {
        // Arrange
        var expectedMarginDecimal = (decimal)expectedMargin;

        // Act & Assert
        shape.TextBox.TopMargin.Should().BeApproximately(expectedMarginDecimal, 0.01m);
    }

    [Test]
    [SlideShape("autoshape-case003.pptx", 1, "AutoShape 2", 3.69)]
    public void BottomMargin_getter_returns_bottom_margin_of_text_frame(IShape shape, double expectedMargin)
    {
        // Arrange
        var expectedMarginDecimal = (decimal)expectedMargin;

        // Act & Assert
        shape.TextBox.BottomMargin.Should().Be(expectedMarginDecimal);
    }

    [Test]
    [TestCase("001.pptx", 1, "TextBox 4")]
    public void VerticalAlignment_Setter_updates_text_vertical_alignment(string presName, int slideNumber,
        string shapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(presName));
        var textbox = pres.Slides[slideNumber - 1].Shapes.Shape<IShape>(shapeName).TextBox!;
        var mStream = new MemoryStream();

        // Act
        textbox.VerticalAlignment = TextVerticalAlignment.Bottom;

        // Assert
        textbox.VerticalAlignment.Should().Be(TextVerticalAlignment.Bottom);

        pres.Save(mStream);
        pres = new Presentation(mStream);
        textbox = pres.Slides[slideNumber - 1].Shapes.Shape<IShape>(shapeName).TextBox!;
        textbox.VerticalAlignment.Should().Be(TextVerticalAlignment.Bottom);
    }

    [Test]
    [TestCase("001.pptx", 1, "TextBox 4")]
    public void TextDirection_Getter_returns_textbox_text_direction(
        string presName,
        int slideNumber,
        string shapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(presName));
        var textbox = pres.Slide(slideNumber).Shape<IShape>(shapeName).TextBox!;

        // Assert
        textbox.TextDirection.Should().Be(TextDirection.Horizontal);
    }

    [Test]
    [TestCase("001.pptx", 2, "Table 5")]
    public void TextDirection_Getter_returns_cell_textbox_text_direction(
        string presName,
        int slideNumber,
        string shapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(presName));
        var textbox = pres.Slide(slideNumber).Shape(shapeName).Table[0, 0].TextBox;

        // Act & Assert
        textbox.TextDirection.Should().Be(TextDirection.Rotate90);
    }

    [Test]
    [TestCase("001.pptx", 1, "TextBox 4")]
    public void TextDirection_Setter_sets_textbox_text_direction(
        string presName,
        int slideNumber,
        string shapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(presName));
        var textbox = pres.Slide(slideNumber).Shape<IShape>(shapeName).TextBox!;
        var mStream = new MemoryStream();

        // Act
        textbox.TextDirection = TextDirection.Rotate270;

        // Assert
        textbox.TextDirection.Should().Be(TextDirection.Rotate270);

        pres.Save(mStream);
        pres = new Presentation(mStream);
        pres.Slide(slideNumber).Shape<IShape>(shapeName).TextBox!.TextDirection.Should()
            .Be(TextDirection.Rotate270);
    }

    [Test]
    [TestCase("054_get_shape_xpath.pptx", 1, "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[1]/p:txBody[1]")]
    [TestCase("054_get_shape_xpath.pptx", 2, "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[1]/p:txBody[1]")]
    public void SDKXPath_returns_xpath_of_underlying_txBody_element(string presentationName, int slideNumber,
        string expectedXPath)
    {
        // Arrange
        var pres = new Presentation(TestAsset(presentationName));
        var textFrame = pres.Slides[slideNumber - 1].GetTexts().First();

        // Act
        var sdkXPath = textFrame.SdkXPath;

        // Assert
        sdkXPath.Should().Be(expectedXPath);
    }
}
