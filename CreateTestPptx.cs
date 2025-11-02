// Temporary C# program to create a test PPTX file with a layout that has:
// 1. A background image in the layout
// 2. Footer, date, and slide number placeholders
// This reproduces issue #1156

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

var outputPath = "tests/ShapeCrawler.DevTests/assets/085_layout_with_background_image.pptx";

// Start with an existing simple presentation to get the basic structure
var sourcePath = "tests/ShapeCrawler.DevTests/assets/001.pptx";

// Copy the source file
File.Copy(sourcePath, outputPath, true);

// Open and modify it
using (var presentationDocument = PresentationDocument.Open(outputPath, true))
{
    var presentationPart = presentationDocument.PresentationPart!;
    var slideMasterPart = presentationPart.SlideMasterParts.First();
    var slideLayoutPart = slideMasterPart.SlideLayoutParts.First();
    
    // Add a test image to the package
    var imagePart = slideLayoutPart.AddImagePart(ImagePartType.Jpeg);
    using (var imageStream = new FileStream("tests/ShapeCrawler.DevTests/assets/images/test.jpg", FileMode.Open))
    {
        imagePart.FeedData(imageStream);
    }
    var imageRelId = slideLayoutPart.GetIdOfPart(imagePart);
    
    // Modify the layout to add background image
    var slideLayout = slideLayoutPart.SlideLayout;
    var cSld = slideLayout.CommonSlideData ?? slideLayout.AppendChild(new CommonSlideData());
    
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
    
    // Add footer placeholder to the layout's shape tree
    var shapeTree = cSld.ShapeTree ?? cSld.AppendChild(new ShapeTree());
    
    // Add footer placeholder
    var footerShape = new P.Shape(
        new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = 100U, Name = "Footer Placeholder" },
            new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
            new P.ApplicationNonVisualDrawingProperties(
                new P.PlaceholderShape { Type = PlaceholderValues.Footer }
            )
        ),
        new P.ShapeProperties(),
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
        new P.ShapeProperties(),
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
        new P.ShapeProperties(),
        new P.TextBody(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph()
        )
    );
    shapeTree.AppendChild(dateShape);
    
    slideLayoutPart.SlideLayout.Save();
}

Console.WriteLine($"Created test file: {outputPath}");

