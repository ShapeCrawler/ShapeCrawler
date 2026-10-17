using System.IO.Compression;
using System.Text;
using FluentAssertions;
using ShapeCrawler.Tests.Helpers;

namespace ShapeCrawler.Tests;

public class PresentationOpeningTests : SCTest
{
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
}
