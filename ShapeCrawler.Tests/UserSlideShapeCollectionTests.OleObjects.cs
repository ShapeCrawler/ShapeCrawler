using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using W = DocumentFormat.OpenXml.Wordprocessing;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Tests;

public partial class UserSlideShapeCollectionTests
{
    [Test]
    public void AddOleObject_embeds_Word_and_Excel_documents_on_first_slide()
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide().Slide());
        using var document = OleDocument(OleObjectType.WordDocument);
        using var workbook = OleDocument(OleObjectType.ExcelWorkbook);
        using var documentPreview = TestAsset("09 png image.png");
        using var workbookPreview = TestAsset("09 png image.png");
        var shapes = presentation.Slide(1).Shapes;

        // Act
        shapes.AddOleObject(50, 60, 200, 100, document, OleObjectType.WordDocument, documentPreview);
        shapes.AddOleObject(300, 60, 200, 100, workbook, OleObjectType.ExcelWorkbook, workbookPreview);

        // Assert
        using var reopened = SaveAndOpenPresentation(presentation);
        reopened.Slide(1).Shapes.Should().HaveCount(2);
        reopened.Slide(2).Shapes.Should().BeEmpty();
        using var sdk = SaveAndOpenPresentationAsSdk(reopened);
        var slideId = sdk.PresentationPart!.Presentation.SlideIdList!.Elements<P.SlideId>().First();
        var slidePart = (SlidePart)sdk.PresentationPart.GetPartById(slideId.RelationshipId!);
        var objects = slidePart.Slide.Descendants<P.OleObject>().ToArray();
        objects.Select(o => o.ProgId!.Value).Should().Equal("Word.Document.12", "Excel.Sheet.12");
        using var embeddedDocument = new MemoryStream(CopyTestPartBytes(slidePart.GetPartById(objects[0].Id!)));
        using var word = WordprocessingDocument.Open(embeddedDocument, false);
        word.MainDocumentPart!.Document.Body!.InnerText.Should().Be("Embedded document");
        using var embeddedWorkbook = new MemoryStream(CopyTestPartBytes(slidePart.GetPartById(objects[1].Id!)));
        using var excel = SpreadsheetDocument.Open(embeddedWorkbook, false);
        excel.WorkbookPart!.Workbook.Sheets!.Elements<X.Sheet>().Single().Name!.Value.Should().Be("Sheet1");
        slidePart.Slide.Descendants<P.NonVisualDrawingProperties>()
            .Select(p => p.Id!.Value).Should().OnlyHaveUniqueItems();
        ValidatePresentation(reopened);
    }

    [TestCase(OleObjectType.WordDocument, "Word.Document.12", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")]
    [TestCase(OleObjectType.ExcelWorkbook, "Excel.Sheet.12", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")]
    [TestCase(OleObjectType.PowerPointPresentation, "PowerPoint.Show.12", "application/vnd.openxmlformats-officedocument.presentationml.presentation")]
    public void AddOleObject_embeds_document_and_preview(
        OleObjectType type, string progId, string contentType)
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());
        var shapes = presentation.Slide(1).Shapes;
        using var content = OleDocument(type);
        var expectedBytes = content.ToArray();
        using var preview = TestAsset("09 png image.png");

        // Act
        shapes.AddOleObject(50, 60, 200, 100, content, type, preview);

        // Assert
        content.CanRead.Should().BeTrue();
        preview.CanRead.Should().BeTrue();
        using var reopened = SaveAndOpenPresentation(presentation);
        var shape = reopened.Slide(1).Shapes.Single();
        shape.X.Should().Be(50);
        shape.Y.Should().Be(60);
        shape.Width.Should().Be(200);
        shape.Height.Should().Be(100);
        using var sdk = SaveAndOpenPresentationAsSdk(reopened);
        var slidePart = sdk.PresentationPart!.SlideParts.Single();
        var ole = slidePart.Slide.Descendants<P.OleObject>().Single();
        ole.ProgId!.Value.Should().Be(progId);
        ole.GetFirstChild<P.OleObjectEmbed>().Should().NotBeNull();
        ole.ImageWidth!.Value.Should().Be(200 * 12700);
        ole.ImageHeight!.Value.Should().Be(100 * 12700);
        var part = slidePart.GetPartById(ole.Id!);
        part.Should().BeOfType<EmbeddedPackagePart>();
        part.ContentType.Should().Be(contentType);
        CopyTestPartBytes(part).Should().Equal(expectedBytes);
        var imageId = ole.Descendants<A.Blip>().Single().Embed!.Value!;
        var imagePart = slidePart.GetPartById(imageId);
        imagePart.ContentType.Should().Be("image/png");
        using var image = new ImageMagick.MagickImage(CopyTestPartBytes(imagePart));
        using var expectedImage = new ImageMagick.MagickImage(preview.ToArray());
        image.Width.Should().Be(expectedImage.Width);
        image.Height.Should().Be(expectedImage.Height);
        image.Compare(expectedImage, ImageMagick.ErrorMetric.Absolute).Should().Be(0);
        ValidatePresentation(reopened);
    }

    [Test]
    public void AddOleObject_assigns_unique_ids_and_supports_copying_to_another_presentation()
    {
        // Arrange
        using var source = new Presentation(p => p.Slide());
        using var target = new Presentation(p => p.Slide());
        using var content = OleDocument(OleObjectType.WordDocument);
        var expectedBytes = content.ToArray();
        using var preview = TestAsset("09 png image.png");

        // Act
        source.Slide(1).Shapes.AddOleObject(20, 30, 100, 80, content, OleObjectType.WordDocument, preview);
        content.Position = 0;
        preview.Position = 0;
        source.Slide(1).Shapes.AddOleObject(120, 30, 100, 80, content, OleObjectType.WordDocument, preview);
        target.Slide(1).Shapes.Add(source.Slide(1).Shapes.Last());
        target.Slide(1).Shapes.Add(source.Slide(1).Shapes.Last());

        // Assert
        using var sourceSdk = SaveAndOpenPresentationAsSdk(source);
        sourceSdk.PresentationPart!.SlideParts.Single().Slide.Descendants<P.NonVisualDrawingProperties>()
            .Select(p => p.Id!.Value).Should().OnlyHaveUniqueItems();
        ValidatePresentation(source);
        using var reopened = SaveAndOpenPresentation(target);
        reopened.Slide(1).Shapes.Should().HaveCount(2);
        using var sdk = SaveAndOpenPresentationAsSdk(reopened);
        var slide = sdk.PresentationPart!.SlideParts.Single();
        slide.Slide.Descendants<P.NonVisualDrawingProperties>().Select(p => p.Id!.Value).Should().OnlyHaveUniqueItems();
        foreach (var ole in slide.Slide.Descendants<P.OleObject>())
        {
            CopyTestPartBytes(slide.GetPartById(ole.Id!)).Should().Equal(expectedBytes);
            slide.GetPartById(ole.Descendants<A.Blip>().Single().Embed!).Should().BeOfType<ImagePart>();
        }

        ValidatePresentation(reopened);
    }

    [TestCase(0, 100, OleObjectType.WordDocument)]
    [TestCase(100, -1, OleObjectType.WordDocument)]
    [TestCase(100, 100, (OleObjectType)999)]
    public void AddOleObject_rejects_invalid_arguments_without_modifying_slide(int width, int height, OleObjectType type)
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());
        using var content = OleDocument(OleObjectType.WordDocument);
        using var preview = TestAsset("09 png image.png");

        // Act
        var add = () => presentation.Slide(1).Shapes.AddOleObject(0, 0, width, height, content, type, preview);

        // Assert
        add.Should().Throw<ArgumentOutOfRangeException>();
        presentation.Slide(1).Shapes.Should().BeEmpty();
        using var sdk = SaveAndOpenPresentationAsSdk(presentation);
        sdk.PresentationPart!.SlideParts.Single().EmbeddedPackageParts.Should().BeEmpty();
    }

    private static MemoryStream OleDocument(OleObjectType type)
    {
        var stream = new MemoryStream();
        switch (type)
        {
            case OleObjectType.WordDocument:
                using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
                {
                    document.AddMainDocumentPart().Document = new W.Document(
                        new W.Body(new W.Paragraph(new W.Run(new W.Text("Embedded document")))));
                }

                break;
            case OleObjectType.ExcelWorkbook:
                using (var workbook = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
                {
                    var part = workbook.AddWorkbookPart();
                    var worksheet = part.AddNewPart<WorksheetPart>();
                    worksheet.Worksheet = new X.Worksheet(new X.SheetData());
                    part.Workbook = new X.Workbook(new X.Sheets(new X.Sheet
                    {
                        Id = part.GetIdOfPart(worksheet), SheetId = 1, Name = "Sheet1"
                    }));
                }

                break;
            default:
                using (var presentation = new Presentation(p => p.Slide()))
                {
                    presentation.Save(stream);
                }

                break;
        }

        stream.Position = 0;
        return stream;
    }
}
