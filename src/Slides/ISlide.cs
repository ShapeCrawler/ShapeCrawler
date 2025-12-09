using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Drawing;
using ShapeCrawler.Presentations;
using ShapeCrawler.Shapes;
using ShapeCrawler.Slides;
using ShapeCrawler.Units;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

#if DEBUG
using System.Threading.Tasks;
#endif

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a slide.
/// </summary>
public interface ISlide
{
    /// <summary>
    ///     Gets or sets custom data. Returns <see langword="null"/> if the custom data is not presented.
    /// </summary>
    string? CustomData { get; set; }

    /// <summary>
    ///     Gets slide layout.
    /// </summary>
    ISlideLayout SlideLayout { get; }

    /// <summary>
    ///     Gets or sets slide number.
    /// </summary>
    int Number { get; set; }

    /// <summary>
    ///     Gets the shape collection.
    /// </summary>
    ISlideShapeCollection Shapes { get; }

    /// <summary>
    ///     Gets the slide notes.
    /// </summary>
    ITextBox? Notes { get; }

    /// <summary>
    ///     Gets the slide fill.
    /// </summary>
    IShapeFill Fill { get; }

    /// <summary>
    ///     Gets all slide text boxes.
    /// </summary>
    public IList<ITextBox> GetTextBoxes();

    /// <summary>
    ///     Hides slide.
    /// </summary>
    void Hide();

    /// <summary>
    ///     Gets a value indicating whether the slide is hidden.
    /// </summary>
    bool Hidden();

    /// <summary>
    ///     Adds specified lines to the slide notes.
    /// </summary>
    void AddNotes(IEnumerable<string> lines);

    /// <summary>
    ///     Gets element by name.
    /// </summary>
    /// <param name="name">element name.</param>
    IShape Shape(string name);

    /// <summary>
    ///     Gets element by ID.
    /// </summary>
    IShape Shape(int id);

    /// <summary>
    ///     Gets shape by name.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T Shape<T>(string name)
        where T : IShape;

    /// <summary>
    ///     Removes the slide.
    /// </summary>
    void Remove();

    /// <summary>
    ///     Saves the slide image to the specified stream.
    /// </summary>
    void SaveImageTo(Stream stream);

#if DEBUG
    /// <summary>
    ///     Saves the slide image to the specified file.
    /// </summary>
    void SaveImageTo(string file);
#endif

    /// <summary>
    ///     Gets a copy of the underlying parent <see cref="PresentationPart"/>.
    /// </summary>
    // ReSharper disable once InconsistentNaming
    PresentationPart GetSDKPresentationPart();

    /// <summary>
    ///     Gets the first shape in the slide.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T First<T>();
}

internal class Slide(ISlideLayout slideLayout, SlideShapeCollection shapes, SlidePart slidePart) : ISlide
{
    private const double Epsilon = 1e-6;

    private static readonly Dictionary<string, Func<A.ColorScheme, A.Color2Type?>> SchemeColorSelectors =
        new(StringComparer.Ordinal)
        {
            { "dk1", scheme => scheme.Dark1Color },
            { "lt1", scheme => scheme.Light1Color },
            { "dk2", scheme => scheme.Dark2Color },
            { "lt2", scheme => scheme.Light2Color },
            { "accent1", scheme => scheme.Accent1Color },
            { "accent2", scheme => scheme.Accent2Color },
            { "accent3", scheme => scheme.Accent3Color },
            { "accent4", scheme => scheme.Accent4Color },
            { "accent5", scheme => scheme.Accent5Color },
            { "accent6", scheme => scheme.Accent6Color },
            { "hlink", scheme => scheme.Hyperlink },
            { "folHlink", scheme => scheme.FollowedHyperlinkColor }
        };

    private readonly TextDrawing textDrawing = new(ParseHexColor);
    private IShapeFill? fill;

    public ISlideLayout SlideLayout => slideLayout;

    public ISlideShapeCollection Shapes => shapes;

    public int Number
    {
        get
        {
            var presDocument = (PresentationDocument)slidePart.OpenXmlPackage;
            var presPart = presDocument.PresentationPart!;
            var currentSlidePartId = presPart.GetIdOfPart(slidePart);
            var slideIdList =
                presPart.Presentation.SlideIdList!.ChildElements.OfType<SlideId>().ToList();
            for (var i = 0; i < slideIdList.Count; i++)
            {
                if (slideIdList[i].RelationshipId == currentSlidePartId)
                {
                    return i + 1;
                }
            }

            throw new SCException("An error occurred while parsing slide number.");
        }

        set
        {
            if (this.Number == value)
            {
                throw new SCException("Slide number is already set to the specified value.");
            }

            var currentIndex = this.Number - 1;
            var newIndex = value - 1;
            var presDocument = (PresentationDocument)slidePart.OpenXmlPackage;
            if (newIndex < 0 || newIndex >= presDocument.PresentationPart!.SlideParts.Count())
            {
                throw new SCException("Slide number is out of range.");
            }

            var presentationPart = presDocument.PresentationPart!;
            var presentation = presentationPart.Presentation;
            var slideIdList = presentation.SlideIdList!;

            // Get the slide ID of the source slide.
            var sourceSlide = (SlideId)slideIdList.ChildElements[currentIndex];

            SlideId? targetSlide;

            // Identify the position of the target slide after which to move the source slide
            if (newIndex == 0)
            {
                targetSlide = null;
            }
            else if (currentIndex < newIndex)
            {
                targetSlide = (SlideId)slideIdList.ChildElements[newIndex];
            }
            else
            {
                targetSlide = (SlideId)slideIdList.ChildElements[newIndex - 1];
            }

            // Remove the source slide from its current position.
            sourceSlide.Remove();
            slideIdList.InsertAfter(sourceSlide, targetSlide);

            presentation.Save();
        }
    }

    public string? CustomData
    {
        get => this.GetCustomData();
        set => this.SetCustomData(value);
    }

    public ITextBox? Notes => this.GetNotes();

    public IShapeFill Fill
    {
        get
        {
            if (this.fill is null)
            {
                var pcSld = slidePart.Slide.CommonSlideData
                            ?? slidePart.Slide.AppendChild<CommonSlideData>(
                                new());

                // Background element needs to be first, else it gets ignored.
                var pBg = pcSld.GetFirstChild<Background>()
                          ?? pcSld.InsertAt<Background>(new(), 0);

                var pBgPr = pBg.GetFirstChild<P.BackgroundProperties>()
                            ?? pBg.AppendChild<BackgroundProperties>(new());

                this.fill = new ShapeFill(pBgPr);
            }

            return this.fill!;
        }
    }

    public bool Hidden() => slidePart.Slide.Show is not null && !slidePart.Slide.Show.Value;

    public void Hide()
    {
        if (slidePart.Slide.Show is null)
        {
            var showAttribute = new OpenXmlAttribute("show", string.Empty, "0");
            slidePart.Slide.SetAttribute(showAttribute);
        }
        else
        {
            slidePart.Slide.Show = false;
        }
    }

    public IShape Shape(string name) => this.Shapes.Shape<IShape>(name);

    public IShape Shape(int id) => this.Shapes.GetById<IShape>(id);

    public T Shape<T>(string name)
        where T : IShape
        => this.Shapes.Shape<T>(name);

    public void SaveImageTo(string file)
    {
        using var fileStream = File.Create(file);
        this.SaveImageTo(fileStream);
    }

    public PresentationPart GetSDKPresentationPart()
    {
        var presDocument = (PresentationDocument)slidePart.OpenXmlPackage;

        return presDocument.Clone().PresentationPart!;
    }

    public T First<T>() => (T)this.Shapes.First(shape => shape is T);

    public IList<ITextBox> GetTextBoxes()
    {
        var collectedTextBoxes = new List<ITextBox>();

        foreach (var shape in this.Shapes)
        {
            this.CollectTextBoxes(shape, collectedTextBoxes);
        }

        return collectedTextBoxes;
    }

    /// <inheritdoc/>
    public void AddNotes(IEnumerable<string> lines)
    {
        var notes = this.Notes;
        if (notes is null)
        {
            this.AddNotesSlide(lines);
        }
        else
        {
            var paragraphs = notes.Paragraphs;
            foreach (var line in lines)
            {
                paragraphs.Add();
                paragraphs[paragraphs.Count - 1].Text = line;
            }
        }
    }

    public void Remove()
    {
        var presDocument = (PresentationDocument)slidePart.OpenXmlPackage;
        var presPart = presDocument.PresentationPart!;
        var pPresentation = presDocument.PresentationPart!.Presentation;
        var slideIdList = pPresentation.SlideIdList!;

        // Find the exact SlideId corresponding to this slide
        var slideIdRelationship = presPart.GetIdOfPart(slidePart);
        var removingPSlideId = slideIdList.Elements<P.SlideId>()
                                   .FirstOrDefault(slideId => slideId.RelationshipId!.Value == slideIdRelationship) ??
                               throw new SCException("Could not find slide ID in presentation.");

        var sectionList = pPresentation.PresentationExtensionList?.Descendants<P14.SectionList>().FirstOrDefault();
        var removingSectionSlideIdListEntry = sectionList?.Descendants<P14.SectionSlideIdListEntry>()
            .FirstOrDefault(s => s.Id! == removingPSlideId.Id!);
        removingSectionSlideIdListEntry?.Remove();

        slideIdList.RemoveChild(removingPSlideId);
        pPresentation.Save();

        var removingSlideIdRelationshipId = removingPSlideId.RelationshipId!;
        new SCPPresentation(pPresentation).RemoveSlideIdFromCustomShow(removingSlideIdRelationshipId.Value!);

        var removingSlidePart = (SlidePart)presPart.GetPartById(removingSlideIdRelationshipId!);
        presPart.DeletePart(removingSlidePart);

        presPart.Presentation.Save();
    }

    public void SaveImageTo(Stream stream)
    {
        this.Save(stream, SKEncodedImageFormat.Png);

        if (stream.CanSeek)
        {
            stream.Position = 0;
        }
    }

    internal void Save(Stream stream, SKEncodedImageFormat format)
    {
        var presPart = this.GetSDKPresentationPart();
        var pSlideSize = presPart.Presentation.SlideSize!;
        var width = new Emus(pSlideSize.Cx!.Value).AsPixels();
        var height = new Emus(pSlideSize.Cy!.Value).AsPixels();

        using var surface = SKSurface.Create(new SKImageInfo((int)width, (int)height));
        var canvas = surface.Canvas;

        this.RenderBackground(canvas);
        this.RenderShapes(canvas);

        using var image = surface.Snapshot();
        using var data = image.Encode(format, 100);
        data.SaveTo(stream);
    }

    private static SKColor ApplyShade(SKColor color, int shadeValue)
    {
        var shadeFactor = shadeValue / 100_000f;

        return new SKColor(
            (byte)(color.Red * shadeFactor),
            (byte)(color.Green * shadeFactor),
            (byte)(color.Blue * shadeFactor),
            color.Alpha);
    }

    private static decimal GetStyleOutlineWidth(IShape shape)
    {
        if (shape.SDKOpenXmlElement is not P.Shape pShape)
        {
            return 0;
        }

        var style = pShape.ShapeStyle;
        var lineRef = style?.LineReference;
        if (lineRef?.Index is null || lineRef.Index.Value == 0)
        {
            return 0;
        }

        // Default line width based on index (idx="2" typically means ~1.5pt line)
        // This is a simplification - proper implementation would look up theme line styles
        var defaultWidth = lineRef.Index.Value * 0.75m;

        return new Points(defaultWidth).AsPixels();
    }

    private static void ApplyRotation(
        SKCanvas canvas,
        IShape shape,
        decimal x,
        decimal y,
        decimal width,
        decimal height)
    {
        if (Math.Abs(shape.Rotation) > Epsilon)
        {
            var centerX = x + (width / 2);
            var centerY = y + (height / 2);
            canvas.RotateDegrees(
                (float)shape.Rotation,
                (float)new Points(centerX).AsPixels(),
                (float)new Points(centerY).AsPixels()
            );
        }
    }

    private static decimal GetShapeOutlineWidth(IShape shape)
    {
        var shapeOutline = shape.Outline;

        // Check for explicit outline weight first
        if (shapeOutline is not null && shapeOutline.Weight > 0)
        {
            return new Points(shapeOutline.Weight).AsPixels();
        }

        // Check for style-based outline width
        var styleWidth = GetStyleOutlineWidth(shape);
        return styleWidth;
    }

    private static SKColor ParseHexColor(string hex, double alphaPercentage)
    {
        hex = hex.TrimStart('#');

        byte r;
        byte g;
        byte b;
        byte a = (byte)(alphaPercentage / 100.0 * 255);

        if (hex.Length == 6)
        {
            r = Convert.ToByte(hex[..2], 16);
            g = Convert.ToByte(hex.Substring(2, 2), 16);
            b = Convert.ToByte(hex.Substring(4, 2), 16);
        }
        else if (hex.Length == 8)
        {
            a = Convert.ToByte(hex[..2], 16);
            r = Convert.ToByte(hex.Substring(2, 2), 16);
            g = Convert.ToByte(hex.Substring(4, 2), 16);
            b = Convert.ToByte(hex.Substring(6, 2), 16);
        }
        else
        {
            return SKColors.Transparent;
        }

        return new SKColor(r, g, b, a);
    }

    private static string? GetHexFromColorElement(A.Color2Type colorElement)
    {
        var rgbColor = colorElement.RgbColorModelHex;
        if (rgbColor?.Val?.Value is { } rgb)
        {
            return rgb;
        }

        var sysColor = colorElement.SystemColor;
        return sysColor?.LastColor?.Value;
    }

    private void RenderRectangle(SKCanvas canvas, IShape shape)
    {
        var x = new Points(shape.X).AsPixels();
        var y = new Points(shape.Y).AsPixels();
        var width = new Points(shape.Width).AsPixels();
        var height = new Points(shape.Height).AsPixels();
        var rect = new SKRect((float)x, (float)y, (float)(x + width), (float)(y + height));

        var cornerRadius = 0m;
        if (shape.GeometryType == Geometry.RoundedRectangle)
        {
            // CornerSize is percentage (0-100), where 100 = half of shortest side
            var shortestSide = Math.Min(width, height);
            cornerRadius = shape.CornerSize / 100m * (shortestSide / 2m);
        }

        canvas.Save();
        ApplyRotation(canvas, shape, shape.X, shape.Y, shape.Width, shape.Height);

        RenderFill(canvas, shape, rect, cornerRadius);
        RenderOutline(canvas, shape, rect, cornerRadius);

        canvas.Restore();
    }

    private void RenderEllipse(SKCanvas canvas, IShape shape)
    {
        var x = new Points(shape.X).AsPixels();
        var y = new Points(shape.Y).AsPixels();
        var width = new Points(shape.Width).AsPixels();
        var height = new Points(shape.Height).AsPixels();
        var rect = new SKRect((float)x, (float)y, (float)(x + width), (float)(y + height));

        canvas.Save();
        ApplyRotation(canvas, shape, shape.X, shape.Y, shape.Width, shape.Height);

        this.RenderEllipseFill(canvas, shape, rect);
        this.RenderEllipseOutline(canvas, shape, rect);

        canvas.Restore();
    }

    private void RenderFill(SKCanvas canvas, IShape shape, SKRect rect, decimal cornerRadius)
    {
        var fillColor = this.GetShapeFillColor(shape);
        if (fillColor is null)
        {
            return;
        }

        using var fillPaint = new SKPaint();
        fillPaint.Color = fillColor.Value;
        fillPaint.Style = SKPaintStyle.Fill;
        fillPaint.IsAntialias = true;

        if (cornerRadius > 0)
        {
            canvas.DrawRoundRect(rect, (float)cornerRadius, (float)cornerRadius, fillPaint);
        }
        else
        {
            canvas.DrawRect(rect, fillPaint);
        }
    }

    private void RenderOutline(SKCanvas canvas, IShape shape, SKRect rect, decimal cornerRadius)
    {
        var outlineColor = this.GetShapeOutlineColor(shape);
        var strokeWidth = GetShapeOutlineWidth(shape);

        if (outlineColor is null || strokeWidth <= 0)
        {
            return;
        }

        using var outlinePaint = new SKPaint();
        outlinePaint.Color = outlineColor.Value;
        outlinePaint.Style = SKPaintStyle.Stroke;
        outlinePaint.StrokeWidth = (float)strokeWidth;
        outlinePaint.IsAntialias = true;

        if (cornerRadius > 0)
        {
            canvas.DrawRoundRect(rect, (float)cornerRadius, (float)cornerRadius, outlinePaint);
        }
        else
        {
            canvas.DrawRect(rect, outlinePaint);
        }
    }

    private void RenderEllipseFill(SKCanvas canvas, IShape shape, SKRect rect)
    {
        var fillColor = this.GetShapeFillColor(shape);
        if (fillColor is null)
        {
            return;
        }

        using var fillPaint = new SKPaint();
        fillPaint.Color = fillColor.Value;
        fillPaint.Style = SKPaintStyle.Fill;
        fillPaint.IsAntialias = true;

        canvas.DrawOval(rect, fillPaint);
    }

    private void RenderEllipseOutline(SKCanvas canvas, IShape shape, SKRect rect)
    {
        var outlineColor = GetShapeOutlineColor(shape);
        var strokeWidth = GetShapeOutlineWidth(shape);

        if (outlineColor is null || strokeWidth <= 0)
        {
            return;
        }

        using var outlinePaint = new SKPaint();
        outlinePaint.Color = outlineColor.Value;
        outlinePaint.Style = SKPaintStyle.Stroke;
        outlinePaint.StrokeWidth = (float)strokeWidth;
        outlinePaint.IsAntialias = true;

        canvas.DrawOval(rect, outlinePaint);
    }

    private SKColor? GetShapeOutlineColor(IShape shape)
    {
        var shapeOutline = shape.Outline;

        // Check for explicit outline color first
        if (shapeOutline?.HexColor is not null)
        {
            return ParseHexColor(shapeOutline.HexColor, 100);
        }

        // Check for style-based outline (lnRef with scheme color)
        var styleColor = GetStyleOutlineColor(shape);
        if (styleColor is not null)
        {
            return styleColor;
        }

        return null;
    }

    private SKColor? GetStyleOutlineColor(IShape shape)
    {
        if (shape.SDKOpenXmlElement is not P.Shape { ShapeStyle.LineReference: { } lineRef })
        {
            return null;
        }

        var schemeColor = lineRef.GetFirstChild<A.SchemeColor>();
        if (schemeColor is null)
        {
            return null;
        }

        var schemeColorValue = schemeColor.Val?.InnerText;
        if (schemeColorValue is null)
        {
            return null;
        }

        var hexColor = this.ResolveSchemeColor(schemeColorValue);
        if (hexColor is null)
        {
            return null;
        }

        var baseColor = ParseHexColor(hexColor, 100);
        var shadeValue = schemeColor.GetFirstChild<A.Shade>()?.Val?.Value;

        return shadeValue is null
            ? baseColor
            : ApplyShade(baseColor, shadeValue.Value);
    }

    private SKColor? GetShapeFillColor(IShape shape)
    {
        var shapeFill = shape.Fill;

        // Check for explicit solid fill first
        if (shapeFill is { Type: FillType.Solid, Color: not null })
        {
            return ParseHexColor(shapeFill.Color, shapeFill.Alpha);
        }

        // Check for style-based fill (fillRef with scheme color)
        if (shapeFill is null || shapeFill.Type == FillType.NoFill)
        {
            var styleColor = this.GetStyleFillColor(shape);
            if (styleColor is not null)
            {
                return styleColor;
            }
        }

        return null;
    }

    private SKColor? GetStyleFillColor(IShape shape)
    {
        var pShape = shape.SDKOpenXmlElement as P.Shape;
        if (pShape is null)
        {
            return null;
        }

        var style = pShape.ShapeStyle;
        var fillRef = style?.FillReference;
        if (fillRef is null)
        {
            return null;
        }

        var schemeColor = fillRef.GetFirstChild<A.SchemeColor>();
        if (schemeColor?.Val is null)
        {
            return null;
        }

        var schemeColorValue = schemeColor.Val?.InnerText;
        if (schemeColorValue is null)
        {
            return null;
        }

        var hexColor = this.ResolveSchemeColor(schemeColorValue);

        return hexColor is not null ? ParseHexColor(hexColor, 100) : null;
    }

    private string? ResolveSchemeColor(string schemeColorName)
    {
        var colorScheme = this.GetColorScheme();
        if (colorScheme is null)
        {
            return null;
        }

        if (!SchemeColorSelectors.TryGetValue(schemeColorName, out var selector))
        {
            return null;
        }

        var colorElement = selector(colorScheme);

        return colorElement is null ? null : GetHexFromColorElement(colorElement);
    }

    private A.ColorScheme? GetColorScheme() =>
        slidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme.ThemeElements?.ColorScheme;

    private void CollectTextBoxes(IShape shape, List<ITextBox> buffer)
    {
        if (shape.TextBox is not null)
        {
            buffer.Add(shape.TextBox);
        }

        if (shape.Table is not null)
        {
            foreach (var cell in shape.Table.Rows.SelectMany(row => row.Cells))
            {
                buffer.Add(cell.TextBox);
            }
        }

        if (shape.GroupedShapes is not null)
        {
            foreach (var innerShape in shape.GroupedShapes)
            {
                this.CollectTextBoxes(innerShape, buffer);
            }
        }
    }

    private ITextBox? GetNotes()
    {
        var notesSlidePart = slidePart.NotesSlidePart;

        if (notesSlidePart is null)
        {
            return null;
        }

        var notesShapes = new ShapeCollection(notesSlidePart);
        var notesPlaceholder = notesShapes
            .FirstOrDefault(shape =>
                shape is { PlaceholderType: not null, TextBox: not null, PlaceholderType: PlaceholderType.Text });
        return notesPlaceholder?.TextBox;
    }

    private void AddNotesSlide(IEnumerable<string> lines)
    {
        // Build up the children of the text body element
        var textBodyChildren = new List<OpenXmlElement>() { new BodyProperties(), new ListStyle() };

        // Add in the text lines
        textBodyChildren.AddRange(
            lines
                .Select(line => new DocumentFormat.OpenXml.Drawing.Paragraph(
                    new ParagraphProperties(),
                    new Run(
                        new RunProperties(),
                        new DocumentFormat.OpenXml.Drawing.Text(line)),
                    new EndParagraphRunProperties())));

        // Always add at least one paragraph, even if empty
        if (!lines.Any())
        {
            textBodyChildren.Add(
                new DocumentFormat.OpenXml.Drawing.Paragraph(
                    new EndParagraphRunProperties()));
        }

        // https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-notes-slides
        var rid = new SCOpenXmlPart(slidePart).NextRelationshipId();
        var notesSlidePart1 = slidePart.AddNewPart<NotesSlidePart>(rid);
        var notesSlide = new NotesSlide(
            new CommonSlideData(
                new ShapeTree(
                    new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                        new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties()
                        {
                            Id = (UInt32Value)1U, Name = string.Empty
                        },
                        new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()),
                    new DocumentFormat.OpenXml.Presentation.Shape(
                        new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
                            new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties()
                            {
                                Id = (UInt32Value)2U, Name = "Notes Placeholder 2"
                            },
                            new DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(
                                new ShapeLocks() { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(
                                new PlaceholderShape() { Type = PlaceholderValues.Body })),
                        new DocumentFormat.OpenXml.Presentation.ShapeProperties(),
                        new DocumentFormat.OpenXml.Presentation.TextBody(
                            textBodyChildren)))),
            new ColorMapOverride(new MasterColorMapping()));
        notesSlidePart1.NotesSlide = notesSlide;
    }

    private string? GetCustomData()
    {
        var getCustomXmlPart = this.GetCustomXmlPartOrNull();
        if (getCustomXmlPart == null)
        {
            return null;
        }

        var customXmlPartStream = getCustomXmlPart.GetStream();
        using var customXmlStreamReader = new StreamReader(customXmlPartStream);
        var raw = customXmlStreamReader.ReadToEnd();
        return raw[3..];
    }

    private void SetCustomData(string? value)
    {
        var getCustomXmlPart = this.GetCustomXmlPartOrNull();
        Stream customXmlPartStream;
        if (getCustomXmlPart == null)
        {
            var newSlideCustomXmlPart = slidePart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            customXmlPartStream = newSlideCustomXmlPart.GetStream();
        }
        else
        {
            customXmlPartStream = getCustomXmlPart.GetStream();
        }

        using var customXmlStreamReader = new StreamWriter(customXmlPartStream);
        customXmlStreamReader.Write($"ctd{value}");
    }

    private CustomXmlPart? GetCustomXmlPartOrNull()
    {
        foreach (var customXmlPart in slidePart.CustomXmlParts)
        {
            using var customXmlPartStream = new StreamReader(customXmlPart.GetStream());
            var customXmlPartText = customXmlPartStream.ReadToEnd();
            if (customXmlPartText.StartsWith(
                    "ctd",
                    StringComparison.CurrentCulture))
            {
                return customXmlPart;
            }
        }

        return null;
    }

    private SKColor GetSkColor()
    {
        var hex = this.Fill.Color!.TrimStart('#');

        // Validate hex length before parsing
        if (hex.Length != 6 && hex.Length != 8)
        {
            return SKColors.White; // used by the PowerPoint application as the default background color
        }

        return ParseHexColor(hex, 100);
    }

    private void RenderBackground(SKCanvas canvas)
    {
        var slideFill = this.Fill;
        if (slideFill is { Type: FillType.Solid, Color: not null })
        {
            var skColor = this.GetSkColor();
            canvas.Clear(skColor);
        }
        else if (slideFill is { Type: FillType.Picture, Picture: not null })
        {
            var bytes = slideFill.Picture.AsByteArray();
            using var stream = new MemoryStream(bytes);
            using var bitmap = SKBitmap.Decode(stream);
            var destRect = new SKRect(0, 0, canvas.DeviceClipBounds.Width, canvas.DeviceClipBounds.Height);
            canvas.DrawBitmap(bitmap, destRect);
        }
        else
        {
            // Default to white for unsupported backgrounds
            canvas.Clear(SKColors.White);
        }
    }

    private void RenderShapes(SKCanvas canvas)
    {
        foreach (var shape in this.Shapes)
        {
            if (shape.Hidden)
            {
                continue;
            }

            this.RenderShape(canvas, shape);
        }
    }

    private void RenderShape(SKCanvas canvas, IShape shape)
    {
        var geometryType = shape.GeometryType;

        switch (geometryType)
        {
            case Geometry.Rectangle:
            case Geometry.RoundedRectangle:
                RenderRectangle(canvas, shape);
                break;
            case Geometry.Ellipse:
                this.RenderEllipse(canvas, shape);
                break;
            default:
                this.RenderText(canvas, shape);
                return;
        }

        this.RenderText(canvas, shape);
    }

    private void RenderText(SKCanvas canvas, IShape shape)
    {
        if (shape.TextBox is null)
        {
            return;
        }

        canvas.Save();
        ApplyRotation(canvas, shape, shape.X, shape.Y, shape.Width, shape.Height);
        this.textDrawing.Render(canvas, shape);
        canvas.Restore();
    }
}