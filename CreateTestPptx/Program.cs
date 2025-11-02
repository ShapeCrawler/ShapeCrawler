// Temporary C# program to create a test PPTX file with a layout that has:
// 1. A background image in the layout
// 2. Footer, date, and slide number placeholders
// This reproduces issue #1156

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

var outputPath = "../tests/ShapeCrawler.DevTests/assets/085_layout_with_background_image.pptx";

// Start with an existing simple presentation to get the basic structure
var sourcePath = "../tests/ShapeCrawler.DevTests/assets/002.pptx";

if (!File.Exists(sourcePath))
{
    Console.WriteLine($"Source file not found: {sourcePath}");
    return;
}

// Copy the source file
File.Copy(sourcePath, outputPath, true);

// Open and modify it
using (var presentationDocument = PresentationDocument.Open(outputPath, true))
{
    var presentationPart = presentationDocument.PresentationPart!;
    var slideMasterPart = presentationPart.SlideMasterParts.First();
    var slideLayoutPart = slideMasterPart.SlideLayoutParts.First();
    
    // Create a simple 1x1 pixel JPEG image in memory
    byte[] imageBytes = new byte[] {
        0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0x4A, 0x46, 0x49, 0x46, 0x00, 0x01,
        0x01, 0x01, 0x00, 0x48, 0x00, 0x48, 0x00, 0x00, 0xFF, 0xDB, 0x00, 0x43,
        0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
        0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
        0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
        0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
        0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
        0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xC0, 0x00, 0x0B, 0x08,
        0x00, 0x01, 0x00, 0x01, 0x01, 0x01, 0x11, 0x00, 0xFF, 0xC4, 0x00, 0x14,
        0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xDA, 0x00, 0x08, 0x01, 0x01,
        0x00, 0x00, 0x3F, 0x00, 0x7F, 0xFF, 0xD9
    };
    
    // Add the image to the layout
    var imagePart = slideLayoutPart.AddImagePart(ImagePartType.Jpeg);
    using (var imageStream = new MemoryStream(imageBytes))
    {
        imagePart.FeedData(imageStream);
    }
    var imageRelId = slideLayoutPart.GetIdOfPart(imagePart);
    
    // Modify the layout to add background image
    var slideLayout = slideLayoutPart.SlideLayout;
    var cSld = slideLayout.CommonSlideData ?? slideLayout.AppendChild(new CommonSlideData());
    
    // Remove existing background if any
    var existingBg = cSld.GetFirstChild<P.Background>();
    existingBg?.Remove();
    
    // Add background with image
    var background = new P.Background(
        new P.BackgroundProperties(
            new A.BlipFill(
                new A.Blip { Embed = imageRelId },
                new A.Stretch(new A.FillRectangle())
            )
        )
    );
    
    // Insert background as first child of CommonSlideData
    cSld.InsertAt(background, 0);
    
    // Get or create shape tree
    var shapeTree = cSld.ShapeTree ?? cSld.AppendChild(new ShapeTree());
    
    // Add group shape properties if not present
    if (shapeTree.GetFirstChild<P.GroupShapeProperties>() == null)
    {
        shapeTree.InsertAt(new P.GroupShapeProperties(), 0);
    }
    
    // Add non-visual group shape properties if not present
    if (shapeTree.GetFirstChild<P.NonVisualGroupShapeProperties>() == null)
    {
        shapeTree.InsertAt(new P.NonVisualGroupShapeProperties(
            new P.NonVisualDrawingProperties { Id = 1U, Name = "" },
            new P.NonVisualGroupShapeDrawingProperties(),
            new P.ApplicationNonVisualDrawingProperties()
        ), 0);
    }
    
    // Add footer placeholder
    var footerShape = new P.Shape(
        new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = 100U, Name = "Footer Placeholder" },
            new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
            new P.ApplicationNonVisualDrawingProperties(
                new P.PlaceholderShape { Type = PlaceholderValues.Footer }
            )
        ),
        new P.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = 0L, Y = 6858000L },
                new A.Extents { Cx = 2971800L, Cy = 457200L }
            )
        ),
        new P.TextBody(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph()
        )
    );
    shapeTree.AppendChild(footerShape);
    
    // Add slide number placeholder
    var slideNumShape = new P.Shape(
        new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = 101U, Name = "Slide Number Placeholder" },
            new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
            new P.ApplicationNonVisualDrawingProperties(
                new P.PlaceholderShape { Type = PlaceholderValues.SlideNumber }
            )
        ),
        new P.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = 6553200L, Y = 6858000L },
                new A.Extents { Cx = 2971800L, Cy = 457200L }
            )
        ),
        new P.TextBody(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph()
        )
    );
    shapeTree.AppendChild(slideNumShape);
    
    // Add date/time placeholder
    var dateShape = new P.Shape(
        new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = 102U, Name = "Date Placeholder" },
            new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
            new P.ApplicationNonVisualDrawingProperties(
                new P.PlaceholderShape { Type = PlaceholderValues.DateAndTime }
            )
        ),
        new P.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = 3276600L, Y = 6858000L },
                new A.Extents { Cx = 2971800L, Cy = 457200L }
            )
        ),
        new P.TextBody(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph()
        )
    );
    shapeTree.AppendChild(dateShape);
    
    slideLayoutPart.SlideLayout.Save();
}

Console.WriteLine($"Created test file: {Path.GetFullPath(outputPath)}");
