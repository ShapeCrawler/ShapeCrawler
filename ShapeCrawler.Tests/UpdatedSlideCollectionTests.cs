using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using ShapeCrawler.Tests.Helpers;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tests;

public class UpdatedSlideCollectionTests : SCTest
{
    [Test]
    public void Add_assigns_unique_imported_layout_ids_without_rewriting_existing_layout_ids()
    {
        // Arrange
        using var source = new Presentation(p => p.Slide());
        using var destination = new Presentation(p => p.Slide());
        using var before = new MemoryStream();
        destination.Save(before);
        using var original = PresentationDocument.Open(before, false);
        var originalMaster = original.PresentationPart!.SlideMasterParts.First();
        var originalIds = originalMaster.SlideMaster!.SlideLayoutIdList!
            .Elements<P.SlideLayoutId>().Select(id => id.Id!.Value).ToArray();

        // Act
        destination.Slides.Add(source.Slides[0]);
        destination.Slides.Add(source.Slides[0]);

        // Assert
        using var saved = new MemoryStream();
        destination.Save(saved);
        using var document = PresentationDocument.Open(saved, false);
        var existingMaster = document.PresentationPart!.SlideMasterParts.Single(master => master.Uri == originalMaster.Uri);
        existingMaster.SlideMaster!.SlideLayoutIdList!.Elements<P.SlideLayoutId>()
            .Select(id => id.Id!.Value).Should().Equal(originalIds);
        var allIds = document.PresentationPart.SlideMasterParts
            .SelectMany(master => master.SlideMaster!.SlideLayoutIdList!.Elements<P.SlideLayoutId>())
            .Select(id => id.Id!.Value).ToArray();
        allIds.Should().OnlyHaveUniqueItems();
        ValidatePresentation(destination);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void Add_does_not_allocate_a_copy_of_unrelated_source_media(bool atPosition)
    {
        // Arrange
        const int payloadSize = 8 * 1024 * 1024;
        using var sourceStream = SourceWithUnrelatedMedia(payloadSize);
        using var source = new Presentation(sourceStream);
        using var destination = new Presentation(p => p.Slide());
        var slide = source.Slides[0];
        var allocatedBefore = GC.GetAllocatedBytesForCurrentThread();

        // Act
        if (atPosition)
        {
            destination.Slides.Add(slide, 1);
        }
        else
        {
            destination.Slides.Add(slide);
        }

        // Assert
        var allocated = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;
        allocated.Should().BeLessThan(payloadSize, "importing one slide must not copy unrelated media");
        destination.Slides.Count.Should().Be(2);
        ValidatePresentation(destination);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void Add_preserves_unsaved_edits_and_does_not_modify_source_layouts(bool atPosition)
    {
        // Arrange
        using var source = new Presentation(TestAsset("pictures-case004.pptx"));
        using var destination = new Presentation(p => p.Slide());
        var slide = source.Slides[0];
        var layoutCount = slide.LayoutSlide.MasterSlide.LayoutSlides.Count();
        slide.Shapes.AddTextBox(10, 20, 100, 30, "Unsaved edit");

        // Act
        if (atPosition)
        {
            destination.Slides.Add(slide, 1);
        }
        else
        {
            destination.Slides.Add(slide);
        }

        // Assert
        var copiedSlide = destination.Slides[atPosition ? 0 : 1];
        copiedSlide.Shapes.Last().TextBox.Text.Should().Be("Unsaved edit");
        slide.LayoutSlide.MasterSlide.LayoutSlides.Count().Should().Be(layoutCount);
        using var savedSource = new MemoryStream();
        source.Save(savedSource);
        using var reopenedSource = new Presentation(savedSource);
        reopenedSource.Slides[0].LayoutSlide.MasterSlide.LayoutSlides.Count().Should().Be(layoutCount);
        ValidatePresentation(destination);
        ValidatePresentation(source);
    }

    [TestCase(false)]
    [TestCase(true)]
    [Explicit("Run in a separate test process with DOTNET_GCHeapHardLimit=0x40000000; see tests/slide-import-memory.md.")]
    public void Add_imports_more_than_2500_slides_with_large_unrelated_media(bool atPosition)
    {
        // Arrange
        using var sourceStream = SourceWithUnrelatedMedia(448 * 1024 * 1024);
        using var source = new Presentation(sourceStream);
        var slide = source.Slides[0];

        // Act
        using var saved = ImportAndSaveSlides(slide, atPosition);

        // Assert
        // Release the disposed writer's DOM before the validator loads another DOM
        // under the same heap budget. Import/save and validation are separate phases.
        GC.Collect();
        GC.WaitForPendingFinalizers();
        using var document = PresentationDocument.Open(saved, false);
        document.PresentationPart!.Presentation.SlideIdList!.ChildElements.Count.Should().Be(2502);
        new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(document).Should().BeEmpty();
        TestContext.Out.WriteLine($"Imported 2501 slides; GC heap limit: {GC.GetGCMemoryInfo().TotalAvailableMemoryBytes:N0} bytes.");
    }

    private static MemoryStream ImportAndSaveSlides(IUserSlide slide, bool atPosition)
    {
        using var destination = new Presentation(p => p.Slide());

        // Act
        for (var i = 0; i < 2501; i++)
        {
            if (atPosition)
            {
                destination.Slides.Add(slide, i + 1);
            }
            else
            {
                destination.Slides.Add(slide);
            }
        }

        // Assert
        TestContext.Progress.WriteLine($"Imported 2501 slides (atPosition={atPosition}); saving and validating.");
        destination.Slides.Count.Should().Be(2502);
        var saved = new MemoryStream();
        destination.Save(saved);
        return saved;
    }

    private static MemoryStream SourceWithUnrelatedMedia(int payloadSize)
    {
        using var seed = new Presentation(p =>
        {
            p.Slide();
            p.Slide();
        });
        using var seedStream = new MemoryStream();
        seed.Save(seedStream);
        string mediaPath;
        using (var document = PresentationDocument.Open(seedStream, true))
        {
            var media = document.PresentationPart!.SlideParts.Last().AddImagePart(ImagePartType.Png);
            using var placeholder = new MemoryStream([0]);
            media.FeedData(placeholder);
            mediaPath = media.Uri.OriginalString.TrimStart('/');
        }

        // Highly compressible opaque media keeps the fixture small. Cloning the whole
        // package still inflates it, reproducing pressure on a constrained managed heap.
        var stream = new MemoryStream();
        seedStream.Position = 0;
        using (var sourceArchive = new ZipArchive(seedStream, ZipArchiveMode.Read, leaveOpen: true))
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in sourceArchive.Entries.Where(entry => entry.FullName != mediaPath))
            {
                using var input = entry.Open();
                using var output = archive.CreateEntry(entry.FullName).Open();
                input.CopyTo(output);
            }

            using var media = archive.CreateEntry(mediaPath).Open();
            var buffer = new byte[64 * 1024];
            for (var remaining = payloadSize; remaining > 0; remaining -= buffer.Length)
            {
                media.Write(buffer, 0, Math.Min(remaining, buffer.Length));
            }
        }

        stream.Position = 0;
        return stream;
    }
}
