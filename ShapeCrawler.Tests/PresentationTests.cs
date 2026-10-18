using System.Globalization;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Fixture;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Helpers;
using ShapeCrawler.Presentations;
using OpenMcdf;

namespace ShapeCrawler.Tests;

public class PresentationTests : SCTest
{
    private readonly Fixtures fixtures = new();

    [TestCase(false)]
    [TestCase(true)]
    public void HighlightText_and_HighlightRegex_process_slides_groups_tables_but_exclude_notes_and_masters(bool regex)
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        var slide = pres.Slides[0];
        slide.Shapes.AddTextBox(0, 0, 100, 50, "a TAG b");
        slide.Shapes.AddTextBox(0, 50, 100, 50, "a TAG b");
        var first = slide.Shapes.Last();
        slide.Shapes.AddTextBox(100, 50, 100, 50, "a TAG b");
        slide.Shapes.Group([first, slide.Shapes.Last()]);
        slide.Shapes.AddTable(0, 100, 1, 1);
        slide.Shapes.Last().Table[0, 0].TextBox.SetText("a TAG b");
        slide.AddNotes(["TAG"]);
        var masterBox = pres.MasterSlides[0].Shapes.First().TextBox;
        masterBox.AutofitType = AutofitType.None;
        masterBox.SetText("TAG");
        pres.Slides.Add(1);
        pres.Slides[1].Shapes.AddTextBox(0, 0, 100, 50, "a TAG b");
        var before = pres.Slides.SelectMany(item => item.GetTexts()).Select(box => box.Text).ToArray();

        // Act
        var count = regex ? pres.HighlightRegex(new Regex(@"\bTAG\b"), new Color("FFFF00"))
            : pres.HighlightText("TAG", new Color("FFFF00"));

        // Assert
        count.Should().Be(5);
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var reopened = new Presentation(saved);
        var boxes = reopened.Slides.SelectMany(item => item.GetTexts()).ToArray();
        boxes.Select(box => box.Text).Should().Equal(before);
        boxes.SelectMany(box => box.Paragraphs).SelectMany(paragraph => paragraph.Portions)
            .Where(portion => !portion.TextHighlightColor.IsTransparent).Select(portion => portion.Text)
            .Should().Equal("TAG", "TAG", "TAG", "TAG", "TAG");
        reopened.Slides[0].Notes!.Paragraphs[0].Portions[0].TextHighlightColor.IsTransparent.Should().BeTrue();
        reopened.MasterSlides[0].Shapes.First().TextBox.Paragraphs[0].Portions[0].TextHighlightColor.IsTransparent.Should().BeTrue();
        ValidatePresentation(reopened);
    }

    [Test]
    public void Highlighting_validates_arguments_even_when_the_presentation_has_no_text_boxes()
    {
        // Arrange
        using var pres = new Presentation();

        // Act
        var literal = () => pres.HighlightText(string.Empty, Color.Black);
        var regex = () => pres.HighlightRegex(null!, Color.Black);
        var color = () => pres.HighlightRegex(new Regex("x"), new Color("FFFF00") { Alpha = -1 });

        // Assert
        literal.Should().Throw<ArgumentException>();
        regex.Should().Throw<ArgumentNullException>();
        color.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Test]
    public void ReplaceFont_updates_group_table_and_notes_fonts_without_changing_text_or_formatting()
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        var slide = pres.Slides[0];
        slide.Shapes.AddTextBox(10, 20, 100, 50, "First");
        var first = slide.Shapes.Last();
        slide.Shapes.AddTextBox(110, 20, 100, 50, "Second");
        slide.Shapes.Group([first, slide.Shapes.Last()]);
        slide.Shapes.AddTable(0, 100, 1, 1);
        slide.Shapes.Last().Table[0, 0].TextBox.SetText("Cell");
        slide.AddNotes(["Notes"]);
        var boxes = slide.GetTexts().Append(slide.Notes!).ToArray();
        foreach (var box in boxes)
        {
            var font = box.Paragraphs[0].Portions[0].Font;
            font.LatinName = "Arial";
            font.EastAsianName = "Arial";
            font.IsBold = true;
            font.Size = 18;
        }

        var texts = boxes.Select(box => box.Text).ToArray();

        // Act
        pres.ReplaceFont("Arial", "Verdana");

        // Assert
        boxes.Select(box => box.Text).Should().Equal(texts);
        foreach (var box in boxes)
        {
            var font = box.Paragraphs[0].Portions[0].Font;
            font.LatinName.Should().Be("Verdana");
            font.EastAsianName.Should().Be("Verdana");
            font.IsBold.Should().BeTrue();
            font.Size.Should().Be(18);
        }

        using var saved = new MemoryStream();
        pres.Save(saved);
        using var reopened = new Presentation(saved);
        reopened.Slides[0].GetTexts().Append(reopened.Slides[0].Notes!).Select(box => box.Text).Should().Equal(texts);
        ValidatePresentation(reopened);
    }

    [TestCase("024_chart.pptx")]
    [TestCase("009_table.pptx")]
    [TestCase("autoshape-grouping.pptx")]
    public void ReplaceFont_updates_all_DrawingML_font_slots_and_preserves_package_content(string asset)
    {
        // Arrange
        using var source = new Presentation(TestAsset(asset));
        source.Slides[0].AddNotes(["Notes"]);
        source.Slides[0].Shapes.AddSmartArt(20, 20, 200, 100, SmartArtType.BasicBlockList);
        using var input = new MemoryStream();
        source.Save(input);
        var oldName = "Replacement Test Font";
        var expectedCount = 0;
        using (var sdk = PresentationDocument.Open(input, true))
        {
            foreach (var part in FontTestParts(sdk))
            {
                if (part.RootElement is not { } root)
                {
                    continue;
                }

                foreach (var element in root.Descendants().Where(element => element.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/main"))
                {
                    if (!element.GetAttributes().Any(attribute => attribute.LocalName == "typeface"))
                    {
                        continue;
                    }

                    element.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("typeface", string.Empty, oldName));
                    expectedCount++;
                }
            }
        }

        var binariesBefore = FontTestBinaryParts(input);
        using var pres = new Presentation(input);

        // Act
        pres.ReplaceFont(oldName, "Verdana");

        // Assert
        expectedCount.Should().BeGreaterThan(10);
        using var saved = new MemoryStream();
        pres.Save(saved);
        using (var sdk = PresentationDocument.Open(saved, false))
        {
            var names = FontTestParts(sdk).Where(part => part.RootElement is not null)
                .SelectMany(part => part.RootElement!.Descendants())
                .Where(element => element.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/main")
                .SelectMany(element => element.GetAttributes().Where(attribute => attribute.LocalName == "typeface"))
                .Select(attribute => attribute.Value).ToArray();
            names.Should().NotContain(oldName);
            names.Count(name => name == "Verdana").Should().Be(expectedCount);
        }

        FontTestBinaryParts(saved).Should().BeEquivalentTo(binariesBefore);
        using var reopened = new Presentation(saved);
        ValidatePresentation(reopened);
    }

    [Test]
    public void ReplaceFont_updates_theme_and_script_fonts_without_materializing_theme_references()
    {
        // Arrange
        using var source = new Presentation(p => p.Slide());
        source.Slides[0].Shapes.AddTextBox(0, 0, 100, 50, "Text");
        source.Slides[0].Shapes.Last().TextBox.Paragraphs[0].Portions[0].Font.LatinName = "+mj-lt";
        using var input = new MemoryStream();
        source.Save(input);
        using (var sdk = PresentationDocument.Open(input, true))
        {
            var part = sdk.PresentationPart!.SlideParts.Single();
            var properties = part.Slide.Descendants<A.RunProperties>().Last();
            properties.AddChild(new A.ComplexScriptFont { Typeface = "Arial", CharacterSet = 1 }, true);
            properties.AddChild(new A.SymbolFont { Typeface = "Arial" }, true);
            var paragraph = part.Slide.Descendants<A.Paragraph>().Last();
            paragraph.ParagraphProperties ??= new A.ParagraphProperties();
            paragraph.ParagraphProperties.AddChild(new A.BulletFont { Typeface = "Arial" }, true);
            paragraph.ParagraphProperties.AddChild(new A.CharacterBullet { Char = "•" }, true);
            var major = part.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme.ThemeElements!.FontScheme!.MajorFont!;
            major.LatinFont!.Typeface = "Arial";
            major.Elements<A.SupplementalFont>().First().Typeface = "Arial";
        }

        using var pres = new Presentation(input);
        var font = pres.Slides[0].Shapes.Last().TextBox.Paragraphs[0].Portions[0].Font;
        var originalName = font.LatinName;

        // Act
        pres.ReplaceFont("Arial", "Verdana");

        // Assert
        originalName.Should().Be("Arial");
        font.LatinName.Should().Be("Verdana");
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var result = PresentationDocument.Open(saved, false);
        var slide = result.PresentationPart!.SlideParts.Single().Slide;
        slide.Descendants<A.LatinFont>().Last().Typeface!.Value.Should().Be("+mj-lt");
        slide.Descendants<A.ComplexScriptFont>().Last().Typeface!.Value.Should().Be("Verdana");
        slide.Descendants<A.ComplexScriptFont>().Last().CharacterSet!.Value.Should().Be(1);
        slide.Descendants<A.SymbolFont>().Last().Typeface!.Value.Should().Be("Verdana");
        slide.Descendants<A.BulletFont>().Last().Typeface!.Value.Should().Be("Verdana");
        var theme = result.PresentationPart.SlideMasterParts.First().ThemePart!.Theme;
        theme.Descendants<A.SupplementalFont>().First().Typeface!.Value.Should().Be("Verdana");
        ValidatePresentation(pres);
    }

    [TestCase("Arial", "Verdana", "Verdana", "arial", "Arial Narrow")]
    [TestCase("arial", "Verdana", "Arial", "Verdana", "Arial Narrow")]
    [TestCase("Absent", "Verdana", "Arial", "arial", "Arial Narrow")]
    [TestCase("Arial", "Arial", "Arial", "arial", "Arial Narrow")]
    public void ReplaceFont_matches_complete_names_case_sensitively(string oldName, string newName, string first, string second, string third)
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        foreach (var name in new[] { "Arial", "arial", "Arial Narrow" })
        {
            pres.Slides[0].Shapes.AddTextBox(0, 0, 100, 50, name);
            pres.Slides[0].Shapes.Last().TextBox.Paragraphs[0].Portions[0].Font.LatinName = name;
        }

        // Act
        pres.ReplaceFont(oldName, newName);

        // Assert
        pres.Slides[0].GetTexts().Select(box => box.Paragraphs[0].Portions[0].Font.LatinName)
            .Should().Equal(first, second, third);
        ValidatePresentation(pres);
    }

    [TestCase(null, "Verdana")]
    [TestCase("", "Verdana")]
    [TestCase(" ", "Verdana")]
    [TestCase("Arial", null)]
    [TestCase("Arial", "")]
    [TestCase("Arial", " ")]
    public void ReplaceFont_rejects_invalid_names_before_mutation(string oldName, string newName)
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        pres.Slides[0].Shapes.AddTextBox(0, 0, 100, 50, "Text");
        var font = pres.Slides[0].Shapes.Last().TextBox.Paragraphs[0].Portions[0].Font;
        font.LatinName = "Arial";

        // Act
        var replace = () => pres.ReplaceFont(oldName, newName);

        // Assert
        replace.Should().Throw<ArgumentException>();
        font.LatinName.Should().Be("Arial");
    }

    private static IEnumerable<OpenXmlPart> FontTestParts(OpenXmlPartContainer container)
    {
        var visited = new HashSet<OpenXmlPart>();
        var pending = new Queue<OpenXmlPart>(container.Parts.Select(item => item.OpenXmlPart));
        while (pending.TryDequeue(out var part))
        {
            if (!visited.Add(part))
            {
                continue;
            }

            yield return part;
            foreach (var child in part.Parts)
            {
                pending.Enqueue(child.OpenXmlPart);
            }
        }
    }

    private static Dictionary<string, byte[]> FontTestBinaryParts(Stream stream)
    {
        stream.Position = 0;
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, true);
        return archive.Entries.Where(entry => !entry.FullName.EndsWith(".xml", StringComparison.Ordinal)
            && !entry.FullName.EndsWith(".rels", StringComparison.Ordinal)).ToDictionary(entry => entry.FullName, entry =>
        {
            using var bytes = new MemoryStream();
            using var content = entry.Open();
            content.CopyTo(bytes);
            return bytes.ToArray();
        });
    }

    [Test]
    public void ReplaceText_updates_slides_groups_and_tables_without_changing_notes()
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        var slide = pres.Slides[0];
        slide.Shapes.AddTextBox(0, 0, 300, 100, "{{name}}");
        slide.Shapes.AddTextBox(0, 100, 150, 100, "Group {{name}}");
        var first = slide.Shapes.Last();
        slide.Shapes.AddTextBox(150, 100, 150, 100, "Group {{name}}");
        var second = slide.Shapes.Last();
        slide.Shapes.Group([first, second]);
        slide.Shapes.AddTable(0, 200, 1, 1);
        slide.Shapes.Last().Table.Rows[0].Cells[0].TextBox.SetText("Table {{name}}");
        slide.AddNotes(["Notes {{name}}"]);
        pres.Slides.Add(1);
        pres.Slides[1].Shapes.AddTextBox(0, 0, 300, 100, "Slide 2 {{name}}");

        // Act
        var count = pres.ReplaceText("{{name}}", "Alice");

        // Assert
        count.Should().Be(5);
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var reopened = new Presentation(saved);
        var texts = reopened.Slides.SelectMany(item => item.GetTexts()).Select(textBox => textBox.Text).ToArray();
        texts.Should().Contain("Alice");
        texts.Count(text => text == "Group Alice").Should().Be(2);
        texts.Should().Contain("Table Alice");
        texts.Should().Contain("Slide 2 Alice");
        texts.Should().NotContain(text => text.Contains("{{name}}", StringComparison.Ordinal));
        reopened.Slides[0].Notes!.Text.Should().Contain("Notes {{name}}");
        ValidatePresentation(reopened);
    }

    [Test]
    public void ReplaceText_supports_case_insensitive_template_tokens()
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        pres.Slides[0].Shapes.AddTextBox(0, 0, 300, 100, "{{Name}} {{NAME}}");

        // Act
        var count = pres.ReplaceText("{{name}}", "Alice", ignoreCase: true);

        // Assert
        count.Should().Be(2);
        pres.Slides[0].Shapes.Last().TextBox!.Text.Should().Be("Alice Alice");
    }

    [TestCase(null, "X")]
    [TestCase("", "X")]
    [TestCase("a", null)]
    public void ReplaceText_validates_arguments_even_when_presentation_has_no_slides(string? oldText, string? newText)
    {
        // Arrange
        using var pres = new Presentation();

        // Act
        var replace = () => pres.ReplaceText(oldText!, newText!);

        // Assert
        replace.Should().Throw<ArgumentException>();
    }

    [TestCase(false, "test password")]
    [TestCase(true, "test password")]
    [TestCase(false, "Пароль 🔒 日本語")]
    [TestCase(true, "Пароль 🔒 日本語")]
    public void Save_encrypts_the_entire_presentation(bool toFile, string password)
    {
        // Arrange
        using var source = TestAsset("009_table.pptx");
        var originalBytes = source.ToArray();
        using IPresentation presentation = new Presentation(source);
        presentation.SlideWidth = 777;
        using var destination = new MemoryStream();
        var file = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid()}.pptx");
        try
        {
            // Act
            if (toFile)
            {
                presentation.Save(file, password);
            }
            else
            {
                presentation.Save(destination, password);
            }

            // Assert
            var encryptedBytes = toFile ? File.ReadAllBytes(file) : destination.ToArray();
            using var encrypted = new MemoryStream(encryptedBytes);
            DocumentFormat.OpenXml.Packaging.OpenXmlPackage.IsEncryptedOfficeFile(encrypted).Should().BeTrue();
            using var decrypted = DecryptPresentation(encryptedBytes, password);
            using var reopened = new Presentation(decrypted);
            reopened.SlideWidth.Should().Be(777);
            reopened.Slides.Count.Should().Be(presentation.Slides.Count);
            using var sdk = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(decrypted, false);
            sdk.PresentationPart!.SlideParts.Should().HaveCount(presentation.Slides.Count);
            source.ToArray().Should().Equal(originalBytes);
            destination.CanWrite.Should().BeTrue();
        }
        finally
        {
            File.Delete(file);
        }
    }

    [Test]
    public void Save_uses_fresh_random_keys_and_rejects_the_wrong_password()
    {
        // Arrange
        using var presentation = new Presentation();
        using var first = new MemoryStream();
        using var second = new MemoryStream();

        // Act
        presentation.Save(first, "correct password");
        presentation.Save(second, "correct password");

        // Assert
        first.ToArray().Should().NotEqual(second.ToArray());
        Action decrypt = () => DecryptPresentation(first.ToArray(), "wrong password").Dispose();
        decrypt.Should().Throw<CryptographicException>();
        Action openWithoutPassword = () => new Presentation(first).Dispose();
        openWithoutPassword.Should().Throw<SCException>();
    }

    [TestCase(1)]
    [TestCase(255)]
    public void Save_accepts_password_length_boundaries(int length)
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());
        using var destination = new MemoryStream();
        var password = new string('a', length);

        // Act
        presentation.Save(destination, password);

        // Assert
        using var decrypted = DecryptPresentation(destination.ToArray(), password);
        using var reopened = new Presentation(decrypted);
        reopened.Slides.Should().ContainSingle();
    }

    [Test]
    public void Save_can_export_plain_pptx_after_encrypting()
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());
        using var encrypted = new MemoryStream();
        using var plain = new MemoryStream();

        // Act
        presentation.Save(encrypted, "password");
        presentation.SlideWidth = 888;
        presentation.Save(plain);

        // Assert
        using var reopened = new Presentation(plain);
        reopened.SlideWidth.Should().Be(888);
        using var decrypted = DecryptPresentation(encrypted.ToArray(), "password");
        using var earlier = new Presentation(decrypted);
        earlier.SlideWidth.Should().NotBe(888);
    }

    [TestCase(null)]
    [TestCase("")]
    [TestCase("too long")]
    public void Save_rejects_invalid_password_without_changing_the_destination(string password)
    {
        // Arrange
        using var presentation = new Presentation();
        using var destination = new MemoryStream(new byte[] { 1, 2, 3 });
        var file = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid()}.pptx");
        File.WriteAllBytes(file, new byte[] { 1, 2, 3 });
        if (password == "too long")
        {
            password = new string('a', 256);
        }

        try
        {
            // Act
            Action saveStream = () => presentation.Save(destination, password);
            Action saveFile = () => presentation.Save(file, password);

            // Assert
            saveStream.Should().Throw<ArgumentException>();
            saveFile.Should().Throw<ArgumentException>();
            destination.ToArray().Should().Equal(1, 2, 3);
            File.ReadAllBytes(file).Should().Equal(1, 2, 3);
        }
        finally
        {
            File.Delete(file);
        }
    }

    [Test]
    public void Save_replaces_existing_stream_content_and_validates_destinations()
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());
        using var destination = new MemoryStream();
        destination.SetLength(1024 * 1024);
        destination.Position = 123;
        using var readOnly = new MemoryStream(new byte[1], writable: false);

        // Act
        presentation.Save(destination, "password");
        Action saveReadOnly = () => presentation.Save(readOnly, "password");
        Action saveNull = () => presentation.Save((Stream)null!, "password");

        // Assert
        destination.Length.Should().BeLessThan(1024 * 1024);
        using var decrypted = DecryptPresentation(destination.ToArray(), "password");
        using var reopened = new Presentation(decrypted);
        reopened.Slides.Should().ContainSingle();
        saveReadOnly.Should().Throw<ArgumentException>();
        saveNull.Should().Throw<ArgumentNullException>();
    }

    [TestCase(false)]
    [TestCase(true)]
    public void Constructor_explains_protection_when_opening_an_encrypted_presentation(bool fromFile)
    {
        // Arrange
        // MIT-licensed Open XML SDK fixture from PR #1969 (see the accompanying license).
        using var stream = TestAsset("089 encrypted.pptx");
        var file = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid()}.pptx");
        File.WriteAllBytes(file, stream.ToArray());

        try
        {
            // Act
            Action open = () =>
            {
                using var presentation = fromFile ? new Presentation(file) : new Presentation(stream);
            };

            // Assert
            var exception = open.Should().Throw<SCException>().Which;
            exception.Message.Should().Contain("encrypted or access-restricted")
                .And.Contain("sensitivity label")
                .And.Contain("password protection")
                .And.Contain("PowerPoint")
                .And.Contain("unencrypted PPTX copy");
            exception.InnerException.Should().NotBeNull();
            stream.CanRead.Should().BeTrue();
            using var exclusiveAccess = File.Open(file, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
        finally
        {
            File.Delete(file);
        }
    }

    [TestCase(false)]
    [TestCase(true)]
    public void Constructor_does_not_describe_corrupt_data_as_protected(bool fromFile)
    {
        // Arrange
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("This is not a PPTX file."));
        var file = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid()}.pptx");
        File.WriteAllBytes(file, stream.ToArray());

        try
        {
            // Act
            Action open = () =>
            {
                using var presentation = fromFile ? new Presentation(file) : new Presentation(stream);
            };

            // Assert
            open.Should().Throw<Exception>().Which.Should().NotBeOfType<SCException>();
            stream.CanRead.Should().BeTrue();
        }
        finally
        {
            File.Delete(file);
        }
    }

    [TestCase(false)]
    [TestCase(true)]
    public void Constructor_opens_unencrypted_presentation_with_sensitivity_label_metadata(bool fromFile)
    {
        // Arrange
        using var stream = new MemoryStream();
        using (var original = new Presentation(p => p.Slide()))
        {
            original.Save(stream);
        }

        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.CreateEntry("docMetadata/LabelInfo.xml");
            using var writer = new StreamWriter(entry.Open());
            writer.Write("<labelList xmlns=\"http://schemas.microsoft.com/office/2020/mipLabelMetadata\">" +
                "<label id=\"00000000-0000-0000-0000-000000000001\" enabled=\"1\" name=\"Confidential\"/>" +
                "</labelList>");
        }

        var file = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid()}.pptx");
        File.WriteAllBytes(file, stream.ToArray());

        try
        {
            // Act
            using var presentation = fromFile ? new Presentation(file) : new Presentation(stream);

            // Assert
            presentation.Slides.Should().ContainSingle();
        }
        finally
        {
            File.Delete(file);
        }
    }

    [Test]
    public void SlideWidth_Getter_returns_presentation_Slides_Width()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));

        // Act & Assert
        pres.SlideWidth.Should().Be(720);
    }

    [Test]
    public void SlideWidth_Setter_sets_presentation_Slides_Width()
    {
        // Arrange
        var pptx = TestAsset("009_table.pptx");
        var pres = new Presentation(pptx);

        // Act
        pres.SlideWidth = 1000;

        // Assert
        pres.SlideWidth.Should().Be(1000);
    }

    [Test]
    public void SlideHeight_Getter_returns_presentation_Slides_Height()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));

        // Act & Assert
        pres.SlideHeight.Should().Be(405);
    }

    [Test]
    public void SlideHeight_Setter_sets_presentation_Slides_Height()
    {
        // Arrange
        var pptx = TestAsset("009_table.pptx");
        var pres = new Presentation(pptx);

        // Act
        pres.SlideHeight = 700;

        // Assert
        pres.SlideHeight.Should().Be(700);
    }

    [Test]
    public void Slides_Count_returns_One_When_presentation_contains_one_slide()
    {
        // Arrange
        var pres17 = new Presentation(TestAsset("017.pptx"));
        var pres16 = new Presentation(TestAsset("016.pptx"));
        var pres75 = new Presentation(TestAsset("075.pptx"));

        // Act & Assert
        pres17.Slides.Count.Should().Be(1);
        pres16.Slides.Count.Should().Be(1);
        pres75.Slides.Count.Should().Be(1);
    }

    [Test]
    public void Slides_Count()
    {
        // Arrange
        var pres = new Presentation(TestAsset("007_2 slides.pptx"));
        var removingSlide = pres.Slides[0];
        var slides = pres.Slides;

        // Act
        removingSlide.Remove();

        // Assert
        slides.Count.Should().Be(1);
    }

    [Test]
    public void Slides_Add_adds_specified_slide_at_the_end_of_slide_collection()
    {
        // Arrange
        var sourceSlide = new Presentation(TestAsset("001.pptx")).Slides[0];
        var destPre = new Presentation(TestAsset("002.pptx"));
        var originSlidesCount = destPre.Slides.Count;
        var expectedSlidesCount = ++originSlidesCount;
        MemoryStream savedPre = new();

        // Act
        destPre.Slides.Add(sourceSlide);

        // Assert
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");

        destPre.Save(savedPre);
        destPre = new Presentation(savedPre);
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");
    }

    [Test]
    public void Slides_Add_adds_source_slide_to_empty_presentation()
    {
        // Arrange
        var sourceSlideAtPosition = new Presentation(TestAsset("001.pptx")).Slide(1);
        var sourceSlideAtEnd = new Presentation(TestAsset("001.pptx")).Slide(1);
        var destinationAtPosition = new Presentation();
        var destinationAtEnd = new Presentation();

        // Act
        destinationAtPosition.Slides.Add(sourceSlideAtPosition, 1);
        destinationAtEnd.Slides.Add(sourceSlideAtEnd);

        // Assert
        destinationAtPosition.Slides.Should().ContainSingle();
        destinationAtEnd.Slides.Should().ContainSingle();
        ValidatePresentation(destinationAtPosition);
        ValidatePresentation(destinationAtEnd);
    }

    [Test]
    public void Slides_Add_preserves_layout_relationships()
    {
        // Arrange
        var sourceSlide = new Presentation(TestAsset("037.pptx")).Slide(1);
        var destination = new Presentation(p => p.Slide());

        // Act
        destination.Slides.Add(sourceSlide, 2);

        // Assert
        destination.Slides.Should().HaveCount(2);
        destination.MasterSlides[0].LayoutSlides.Should().Contain(layout => layout.Name == sourceSlide.LayoutSlide.Name);
        ValidatePresentation(destination);
    }

    [Test]
    public void Slides_Add()
    {
        // Arrange
        var pres = new Presentation(p => { p.SlideMaster(1).SlideLayout(1).Background.Picture(fixtures.Image()); });

        // Act
        pres.Slides.Add(1);

        // Assert
        pres.Slides.Count.Should().Be(1);
        ValidatePresentation(pres);
    }

    [Test]
    public void Slides_Add_should_copy_only_layout_of_copying_slide()
    {
        // Arrange
        var sourcePres = new Presentation(TestAsset("pictures-case004.pptx"));
        var copyingSlide = sourcePres.Slides[0];
        var destPres = new Presentation(TestAsset("autoshape-grouping.pptx"));
        var expectedCount = destPres.Slides.Count + 1;
        MemoryStream savedPre = new();

        // Act
        destPres.Slides.Add(copyingSlide);

        // Assert
        destPres.Slides.Count.Should().Be(expectedCount);

        destPres.Save(savedPre);
        destPres = new Presentation(savedPre);
        destPres.Slides.Count.Should().Be(expectedCount);
        destPres.Slides[1].LayoutSlide.MasterSlide.LayoutSlides.Count().Should().Be(1);
        ValidatePresentation(destPres);
    }

    [Test]
    public void Slides_Add_scales_picture_to_fit_target_slide_when_copying_to_presentation_with_different_slide_size()
    {
        // Arrange
        var sourcePres = new Presentation(TestAsset("087 issue-587-source.pptx"));
        var copyingSlide = sourcePres.Slide(1);
        var copyingPicture = copyingSlide.Shape("Picture 4");
        var destPres = new Presentation(TestAsset("088 issue-587-target.pptx"));
        var scaleToFitTargetSize = Math.Min(
            destPres.SlideWidth / sourcePres.SlideWidth,
            destPres.SlideHeight / sourcePres.SlideHeight);
        var targetXOffset = (destPres.SlideWidth - sourcePres.SlideWidth * scaleToFitTargetSize) / 2;
        var targetYOffset = (destPres.SlideHeight - sourcePres.SlideHeight * scaleToFitTargetSize) / 2;
        var expectedPictureX = copyingPicture.X * scaleToFitTargetSize + targetXOffset;
        var expectedPictureY = copyingPicture.Y * scaleToFitTargetSize + targetYOffset;
        var expectedPictureWidth = copyingPicture.Width * scaleToFitTargetSize;
        var expectedPictureHeight = copyingPicture.Height * scaleToFitTargetSize;
        var savedPres = new MemoryStream();

        // Act
        destPres.Slides.Add(copyingSlide);
        destPres.Save(savedPres);
        destPres = new Presentation(savedPres);

        // Assert
        var copiedPicture = destPres.Slides.Last().Shape("Picture 4");
        copiedPicture.X.Should().Be(expectedPictureX);
        copiedPicture.Y.Should().Be(expectedPictureY);
        copiedPicture.Width.Should().Be(expectedPictureWidth);
        copiedPicture.Height.Should().Be(expectedPictureHeight);
        ValidatePresentation(destPres);
    }

    [Test]
    public void Slides_Add_should_copy_notes()
    {
        // Arrange
        var sourcePres = new Presentation(TestAsset("008.pptx"));
        var copyingSlide = sourcePres.Slides[0];
        var destPres = new Presentation(TestAsset("autoshape-case017_slide-number.pptx"));

        // Act
        destPres.Slides.Add(copyingSlide);

        // Assert
        destPres.Slides.Last().Notes!.Text.Should().Be("Test note");
        ValidatePresentation(destPres);
    }

    [Test]
    public void Slides_Add_adds_a_new_slide()
    {
        // Arrange
        var pres = new Presentation();
        var layout = pres.SlideMaster(1).LayoutSlides.First(l => l.Name == "Blank");

        // Act
        pres.Slides.Add(layout.Number);

        // Assert
        pres.Slides.Count.Should().Be(1);
        ValidatePresentation(pres);
    }

    [Test]
    public void Slides_Add_adds_a_new_slide_using_blank_layout()
    {
        // Arrange
        var pres = new Presentation();
        var layout = pres.SlideMaster(1).SlideLayout("Blank");
        var stream = new MemoryStream();

        // Act
        pres.Slides.Add(layout.Number);

        // Assert
        pres.Save(stream);
        new Presentation(stream).Slide(1).Shapes.Should().BeEmpty();
    }

    [Test]
    public void Slides_Add_adds_a_new_slide_using_layout_with_one_shape()
    {
        // Arrange
        var pres = new Presentation();
        var layout = pres.SlideMaster(1).SlideLayout(6);
        var stream = new MemoryStream();

        // Act
        pres.Slides.Add(layout.Number);

        // Assert
        pres.Save(stream);
        new Presentation(stream).Slide(1).Shapes.Count.Should().Be(1);
    }

    [Test]
    public void Slides_Add_should_not_break_hyperlink()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-case018_rotation.pptx"));
        var inserting = pres.Slide(1);

        // Act
        pres.Slides.Add(inserting, 2);

        // Assert
        ValidatePresentation(pres);
    }

    [Test]
    public void Slides_Add_same_presentation_should_not_duplicate_slide_masters()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var sourceSlide = pres.Slide(1);
        var expectedSlideMastersCount = pres.MasterSlides.Count();

        // Act
        pres.Slides.Add(sourceSlide);
        pres.Slides.Add(sourceSlide);

        // Assert
        pres.MasterSlides.Count().Should().Be(expectedSlideMastersCount);
        ValidatePresentation(pres);
    }

    [Test]
    public void SlideMastersCount_ReturnsNumberOfMasterSlidesInThePresentation()
    {
        // Arrange
        var pres1 = new Presentation(TestAsset("001.pptx"));
        var pres2 = new Presentation(TestAsset("002.pptx"));

        // Act
        var slideMastersCountCase1 = pres1.MasterSlides.Count();
        var slideMastersCountCase2 = pres2.MasterSlides.Count();

        // Assert
        slideMastersCountCase1.Should().Be(1);
        slideMastersCountCase2.Should().Be(2);
    }

    [Test]
    public void SlideMaster_Shapes_Count_returns_number_of_master_shapes()
    {
        // Arrange
        var pptx = TestAsset("001.pptx");
        var pres = new Presentation(pptx);

        // Act
        var masterShapesCount = pres.MasterSlides[0].Shapes.Count;

        // Assert
        masterShapesCount.Should().Be(7);
    }

    [Test]
    public void Sections_Remove_removes_specified_section()
    {
        // Arrange
        var pptxStream = TestAsset("autoshape-case017_slide-number.pptx");
        var pres = new Presentation(pptxStream);
        var removingSection = pres.Sections[0];

        // Act
        pres.Sections.Remove(removingSection);

        // Assert
        pres.Sections.Count.Should().Be(0);
    }

    [Test]
    public void Sections_Remove_should_remove_section_after_Removing_Slide_from_section()
    {
        // Arrange
        var pptxStream = TestAsset("autoshape-case017_slide-number.pptx");
        var pres = new Presentation(pptxStream);
        var removingSection = pres.Sections[0];

        // Act
        pres.Slides[0].Remove();
        pres.Sections.Remove(removingSection);

        // Assert
        pres.Sections.Count.Should().Be(0);
    }

    [Test]
    public void Sections_Section_Slides_Count_returns_Zero_When_section_is_Empty()
    {
        // Arrange
        var pptxStream = TestAsset("008.pptx");
        var pres = new Presentation(pptxStream);
        var section = pres.Sections.GetByName("Section 2");

        // Act
        var slidesCount = section.Slides.Count;

        // Assert
        slidesCount.Should().Be(0);
    }

    [Test]
    public void Sections_Section_Slides_Count_returns_number_of_slides_in_section()
    {
        var pptxStream = TestAsset("autoshape-case017_slide-number.pptx");
        var pres = new Presentation(pptxStream);
        var section = pres.Sections.GetByName("Section 1");

        // Act
        var slidesCount = section.Slides.Count;

        // Assert
        slidesCount.Should().Be(1);
    }

    [Test]
    public void Save_saves_presentation_opened_from_Stream_when_it_was_Saved()
    {
        // Arrange
        var presStream = TestAsset("autoshape-case003.pptx");
        var pres = new Presentation(presStream);
        var textBox = pres.Slides[0].Shapes.Shape<IShape>("AutoShape 2").TextBox!;
        textBox.SetText("Test");

        // Act
        pres.Save();

        // Assert
        pres = new Presentation(presStream);
        textBox = pres.Slides[0].Shapes.Shape<IShape>("AutoShape 2").TextBox!;
        textBox.Text.Should().Be("Test");
    }

    [Test]
    public void Save_should_not_throw_exception()
    {
        var presBytes = TestAsset("001.pptx").ToArray();
        var nonExpandableStream = new MemoryStream(presBytes);
        var pres = new Presentation(nonExpandableStream);
        var outputStream = new MemoryStream();

        // Act
        var saving = () => pres.Save(outputStream);

        // Assert
        saving.Should().NotThrow<Exception>();
    }

    [Test]
    public void Save_sets_the_date_of_the_last_modification()
    {
        // Arrange
        var expectedCreated = DateTime.Parse("2024-01-01T12:34:56Z", CultureInfo.InvariantCulture);
        SCSettings.TimeProvider = new FakeTimeProvider(expectedCreated);
        var pres = new Presentation();
        var expectedModified = DateTime.Parse("2024-02-02T15:30:45Z", CultureInfo.InvariantCulture);
        SCSettings.TimeProvider = new FakeTimeProvider(expectedModified);
        var stream = new MemoryStream();

        // Act
        pres.Save(stream);

        // Assert
        stream.Position = 0;
        var updatedPres = new Presentation(stream);
        updatedPres.Properties.Modified.Should().Be(expectedModified);
    }

    [Test]
    public void Footer_AddSlideNumber_adds_slide_number()
    {
        // Arrange
        var pres = new Presentation(pres => { pres.Slide(); });

        // Act
        pres.Footer.AddSlideNumber();

        // Assert
        pres.Footer.SlideNumberAdded().Should().BeTrue();
    }

    [Test]
    public void Footer_RemoveSlideNumber_removes_slide_number()
    {
        // Arrange
        var pres = new Presentation();
        pres.Footer.AddSlideNumber();

        // Act
        pres.Footer.RemoveSlideNumber();

        // Assert
        pres.Footer.SlideNumberAdded().Should().BeFalse();
    }

    [Test]
    public void Footer_SlideNumberAdded_returns_false_When_slide_number_is_not_added()
    {
        // Arrange
        var pres = new Presentation();

        // Act-Assert
        pres.Footer.SlideNumberAdded().Should().BeFalse();
    }

    [Test]
    public void Footer_AddText_adds_text_footers_in_all_slides()
    {
        // Arrange
        var pres = new Presentation(p =>
        {
            p.Slide();
            p.Slide();
        });

        var text = "To infinity and beyond!";

        // Act
        pres.Footer.AddText(text);

        // Assert
        pres.Slides.Should().AllSatisfy(slide =>
        {
            slide.Shapes
                .Should()
                .Contain(shape =>
                    shape.PlaceholderType == PlaceholderType.Footer
                    && shape.TextBox.Text == text);
        });
    }

    [Test]
    public void Footer_RemoveText_removes_text_footers_from_all_slides()
    {
        // Arrange
        var pres = new Presentation(p =>
        {
            p.Slide();
            p.Slide();
        });

        var text = "To infinity and beyond!";

        // Act
        pres.Footer.AddText(text);
        pres.Footer.RemoveText();

        // Assert
        pres.Slides.Should().AllSatisfy(slide =>
        {
            slide.Shapes
                .Should()
                .NotContain(shape => shape.PlaceholderType == PlaceholderType.Footer);
        });
    }

    [Test]
    public void Footer_AddTextOnSlide_adds_text_footer_on_specific_slide()
    {
        // Arrange
        var pres = new Presentation(p =>
        {
            p.Slide();
            p.Slide();
        });

        var text = "To infinity and beyond";

        // Act
        pres.Footer.AddTextOnSlide(2, text);

        var addInOutOfRangeSlide = () => pres.Footer.AddTextOnSlide(3, "Ow snap!");

        // Assert
        pres.Slides[1].Shapes.Should()
            .Contain(s => s.PlaceholderType == PlaceholderType.Footer && s.TextBox.Text == text);
        pres.Slides[0].Shapes.Should().NotContain(s => s.PlaceholderType == PlaceholderType.Footer);

        addInOutOfRangeSlide.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Test]
    public void Footer_RemoveTextOnSlide_removes_text_footer_from_specific_slide()
    {
        // Arrange
        var pres = new Presentation(p =>
        {
            p.Slide();
            p.Slide();
        });
        var text = "To infinity and beyond!";
        pres.Footer.AddText(text);

        // Act
        pres.Footer.RemoveTextOnSlide(1);
        var removeFromOutOfRangeSlide = () => pres.Footer.RemoveTextOnSlide(0);

        // Assert
        pres.Slides[0].Shapes.Should().NotContain(s => s.PlaceholderType == PlaceholderType.Footer);
        pres.Slides[1].Shapes.Should()
            .Contain(s => s.PlaceholderType == PlaceholderType.Footer && s.TextBox.Text == text);

        removeFromOutOfRangeSlide.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Test]
    public void Slides_Add_adds_slide()
    {
        // Arrange
        var sourceSlide = new Presentation(TestAsset("001.pptx")).Slides[0];
        var pptx = TestAsset("002.pptx");
        var destPre = new Presentation(pptx);
        var originSlidesCount = destPre.Slides.Count;
        var expectedSlidesCount = ++originSlidesCount;
        MemoryStream savedPre = new();

        // Act
        destPre.Slides.Add(sourceSlide);

        // Assert
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");

        destPre.Save(savedPre);
        destPre = new Presentation(savedPre);
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");
    }

    [Test]
    public void Add_adds_slide()
    {
        // Arrange
        var pres = new Presentation(TestAsset("017.pptx"));
        var layout = pres.SlideMaster(1).SlideLayout("Title and Content");

        // Act
        pres.Slides.Add(layout.Number);

        // Assert
        pres.Slide(2).Shapes.Count.Should().Be(2);
        ValidatePresentation(pres);
    }

    [Test]
    public void Slides_Add_copies_slide_with_chart()
    {
        // Arrange
        var file = fixtures.AssemblyStream("084 charts.pptx");
        var pres = new Presentation(file);
        var slide = pres.Slides.Last();

        // Act
        pres.Slides.Add(slide, 1);

        // Assert
        ValidatePresentation(pres);
    }


    [Test]
    public void Slides_Add_adds_slide_at_position()
    {
        // Arrange
        var pres = new Presentation(p =>
        {
            p.Slide();
            p.Slide(s =>
            {
                s.Picture(
                    "Picture 1",
                    fixtures.Int(),
                    fixtures.Int(),
                    fixtures.Int(),
                    fixtures.Int(),
                    fixtures.Image());
            });
        });
        var copyingSlide = pres.Slide(2);

        // Act
        pres.Slides.Add(copyingSlide, 1);

        // Assert
        pres.Slide(1).Shapes.Shape("Picture 1").Picture.Image!.AsByteArray().Should().NotBeEmpty();
    }

    [Test]
    [TestCase("007_2 slides.pptx", 1)]
    [TestCase("006_1 slides.pptx", 0)]
    public void Slides_Remove_removes_slide(string file, int expectedSlidesCount)
    {
        // Arrange
        var pres = new Presentation(TestAsset(file));
        var removingSlide = pres.Slides[0];
        var mStream = new MemoryStream();

        // Act
        removingSlide.Remove();

        // Assert
        pres.Slides.Should().HaveCount(expectedSlidesCount);

        pres.Save(mStream);
        pres = new Presentation(mStream);
        pres.Slides.Should().HaveCount(expectedSlidesCount);
    }

    [Test]
    public void Slides_Add_adds_slide_at_the_specified_position()
    {
        // Arrange
        var sourceSlide = new Presentation(TestAsset("001.pptx")).Slide(1);
        var sourceSlideId = Guid.NewGuid().ToString();
        sourceSlide.CustomData = sourceSlideId;
        var destPres = new Presentation(TestAsset("002.pptx"));

        // Act
        destPres.Slides.Add(sourceSlide, 2);

        // Assert
        destPres.Slide(2).CustomData.Should().Be(sourceSlideId);
    }

    [Test]
    public void Slides_Add_adds_a_new_slide_at_the_specified_position_using_specified_layout()
    {
        // Arrange
        var pres = new Presentation(p =>
        {
            p.Slide(s =>
            {
                s.TextShape(Fixtures.String(), fixtures.Int(), fixtures.Int(), fixtures.Int(), fixtures.Int(),
                    Fixtures.String());
            });
        });
        var layoutNumber = pres.MasterSlides.Select(sm => sm.SlideLayout("Blank")).First().Number;

        // Act
        pres.Slides.Add(layoutNumber, 1);

        // Assert
        pres.Slide(2).Shapes.Count.Should().Be(1);
    }


    [Test]
    public void Properties_Title_Setter_sets_title()
    {
        // Arrange
        var pres = new Presentation();
        var expectedCreated = new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Local);

        // Act
        pres.Properties.Title = "Properties_setter_sets_values";
        pres.Properties.Created = expectedCreated;

        // Assert
        pres.Properties.Title.Should().Be("Properties_setter_sets_values");
        pres.Properties.Created.Should().Be(expectedCreated);
    }

    [Test]
    public void Properties_getters_return_valid_values_after_saving_presentation()
    {
        // Arrange
        var pres = new Presentation();
        var expectedCreated = new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Local);
        var stream = new MemoryStream();

        // Act
        pres.Properties.Title = "Properties_setter_survives_round_trip";
        pres.Properties.Created = expectedCreated;
        pres.Properties.RevisionNumber = 100;
        pres.Save(stream);

        // Assert
        stream.Position = 0;
        var updatePres = new Presentation(stream);
        updatePres.Properties.Title.Should().Be("Properties_setter_survives_round_trip");
        updatePres.Properties.Created.Should().Be(expectedCreated);
        pres.Properties.RevisionNumber.Should().Be(100);
    }

    [Test]
    public void Properties_Modified_Getter_returns_date_of_the_last_modification()
    {
        var pres = new Presentation(TestAsset("059_crop-images.pptx"));
        var expectedModified = DateTime.Parse("2024-12-16T17:11:58Z", CultureInfo.InvariantCulture);

        // Act-Assert
        pres.Properties.Modified.Should().Be(expectedModified);
        pres.Properties.Title.Should().Be("");
        pres.Properties.RevisionNumber.Should().Be(7);
        pres.Properties.Comments.Should().BeNull();
    }

    [Test]
    public void Properties_Company()
    {
        // Arrange
        var pres = new Presentation(TestAsset("010.pptx"));

        // Act-Assert
        pres.Properties.Company.Should().BeEquivalentTo("PricewaterhouseCoopers");
    }

    [Test]
    public void Non_parameter_constructor_sets_the_date_of_the_last_modification()
    {
        // Arrange
        var expectedModified = DateTime.Parse("2024-01-01T12:34:56Z", CultureInfo.InvariantCulture);
        SCSettings.TimeProvider = new FakeTimeProvider(expectedModified);

        // Act
        var pres = new Presentation();

        // Assert
        pres.Properties.Modified.Should().Be(expectedModified);
    }

    [Test]
    public void Constructor_does_not_throw_exception_When_the_specified_file_is_a_google_slide_export()
    {
        // Act
        var openingGoogleSlides = () => new Presentation(TestAsset("074 google slides.pptx"));

        // Assert
        openingGoogleSlides.Should().NotThrow();
    }

    [Test]
    public void Create_creates_new_presentation_with_slide()
    {
        // Arrange
        var imageStream = TestAsset("reference image.png");

        // Act
        var pres = new Presentation(pres =>
        {
            pres.Slide(slide =>
            {
                slide.Picture(
                    name: "Picture",
                    x: 100,
                    y: 100,
                    width: 200,
                    height: 50,
                    image: imageStream);
            });
        });

        // Assert
        pres.Slide(1).Shape("Picture").Should().NotBeNull();
    }

    [Test]
    public void Constructor_creates_presentation()
    {
        // Act
        var pres = new Presentation();

        // Assert
        pres.Should().NotBeNull();
        ValidatePresentation(pres);
    }

    [Test]
    public void Constructor_creates_empty_presentation()
    {
        // Act-Assert
        new Presentation().Slides.Should().BeEmpty();
    }

    [Test]
    [Platform(Exclude = "Linux,MacOSX", Reason = "\n vs. \r\n")]
    public void AsMarkdown_returns_markdown_string()
    {
        // Arrange
        var pres = new Presentation(TestAsset("076 bitcoin.pptx"));
        var expectedMarkdown = StringOf("076 bitcoin.md");

        // Act
        var actualMarkdown = pres.AsMarkdown();

        // Assert
        actualMarkdown.Should().BeEquivalentTo(expectedMarkdown);
    }

    [Test]
    public void Save_does_not_throw_exception_When_stream_is_a_File_stream()
    {
        // Arrange
        var pres = new Presentation();
        var file = Path.GetTempFileName();
        using var stream = File.OpenWrite(file);

        // Act
        var saving = () => pres.Save(stream);

        // Assert
        saving.Should().NotThrow();

        // Cleanup
        stream.Close();
        File.Delete(file);
    }

    [Test]
    public void Slides_RemoveThenAdd_EnsuresUniqueSlideIds()
    {
        // Arrange
        var sourceSlide = new Presentation(TestAsset("001.pptx")).Slide(1);
        var destPres = new Presentation(TestAsset("001.pptx"));

        // Act
        destPres.Slide(2).Remove();
        destPres.Slides.Add(sourceSlide, 1);

        // Assert
        using var sdkPresentation = SaveAndOpenPresentationAsSdk(destPres);
        var slideIdRelationshipIdList =
            sdkPresentation.PresentationPart!.Presentation.SlideIdList!.OfType<SlideId>()
                .Select(s => s.RelationshipId);
        slideIdRelationshipIdList.Should().OnlyHaveUniqueItems();
    }

    [Test]
    public void Slide_throws_exception()
    {
        // Arrange
        var pres = new Presentation();

        // Act
        var accessUnavailableSlide = () => pres.Slide(1);

        // Assert
        accessUnavailableSlide.Should().Throw<Exception>();
    }

    private static MemoryStream DecryptPresentation(byte[] bytes, string password)
    {
        using var container = new MemoryStream(bytes);
        using var root = RootStorage.Open(container);
        using var info = root.OpenStream("EncryptionInfo");
        using var reader = new BinaryReader(info);
        reader.ReadUInt16().Should().Be(4);
        reader.ReadUInt16().Should().Be(4);
        reader.ReadUInt32().Should().Be(64);
        var xml = XDocument.Load(info);
        var keyData = xml.Descendants().Single(x => x.Name.LocalName == "keyData");
        var passwordData = xml.Descendants().Single(x => x.Name.LocalName == "encryptedKey");
        keyData.Attribute("keyBits")!.Value.Should().Be("256");
        keyData.Attribute("hashAlgorithm")!.Value.Should().Be("SHA512");
        var salt = Convert.FromBase64String(passwordData.Attribute("saltValue")!.Value);
        byte[] Hash(byte[] first, byte[] second) => SHA512.HashData(first.Concat(second).ToArray());
        byte[] Value(XElement element, string name) => Convert.FromBase64String(element.Attribute(name)!.Value);
        byte[] Decrypt(byte[] data, byte[] key, byte[] iv)
        {
            using var aes = Aes.Create();
            aes.Key = key;
            return aes.DecryptCbc(data, iv, PaddingMode.None);
        }

        var hash = Hash(salt, Encoding.Unicode.GetBytes(password));
        for (var i = 0; i < int.Parse(passwordData.Attribute("spinCount")!.Value, CultureInfo.InvariantCulture); i++)
        {
            var counter = new byte[4];
            BinaryPrimitives.WriteInt32LittleEndian(counter, i);
            hash = Hash(counter, hash);
        }

        byte[] PasswordValue(string attribute, string block) => Decrypt(
            Value(passwordData, attribute), Hash(hash, Convert.FromHexString(block))[..32], salt);
        var verifier = PasswordValue("encryptedVerifierHashInput", "FEA7D2763B4B9E79");
        var expectedHash = PasswordValue("encryptedVerifierHashValue", "D7AA0F6D3061344E");
        if (!CryptographicOperations.FixedTimeEquals(SHA512.HashData(verifier), expectedHash))
        {
            throw new CryptographicException("Invalid password.");
        }

        var key = PasswordValue("encryptedKeyValue", "146E0BE7ABACD0D6");
        var dataSalt = Value(keyData, "saltValue");
        using var payload = root.OpenStream("EncryptedPackage");
        using var encryptedBytes = new MemoryStream();
        payload.CopyTo(encryptedBytes);
        var encrypted = encryptedBytes.ToArray();
        var integrity = xml.Descendants().Single(x => x.Name.LocalName == "dataIntegrity");
        var hmacKey = Decrypt(Value(integrity, "encryptedHmacKey"), key,
            Hash(dataSalt, Convert.FromHexString("5FB2AD010CB9E1F6"))[..16]);
        var hmacValue = Decrypt(Value(integrity, "encryptedHmacValue"), key,
            Hash(dataSalt, Convert.FromHexString("A0677F02B22C8433"))[..16]);
        HMACSHA512.HashData(hmacKey, encrypted).Should().Equal(hmacValue);
        var size = BinaryPrimitives.ReadInt64LittleEndian(encrypted);
        var result = new MemoryStream();
        for (var offset = 8; offset < encrypted.Length; offset += 4096)
        {
            var block = new byte[4];
            BinaryPrimitives.WriteInt32LittleEndian(block, (offset - 8) / 4096);
            var decrypted = Decrypt(encrypted[offset..Math.Min(offset + 4096, encrypted.Length)], key, Hash(dataSalt, block)[..16]);
            result.Write(decrypted, 0, (int)Math.Min(decrypted.Length, size - result.Length));
        }

        result.Position = 0;
        return result;
    }
}
