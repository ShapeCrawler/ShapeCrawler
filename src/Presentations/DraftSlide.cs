using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft for building a slide.
/// </summary>
public sealed class DraftSlide
{
    private readonly List<Action<IUserSlide, Presentation>> actions = [];

    /// <summary>
    ///     Adds a picture to the slide with the specified name and geometry in points.
    /// </summary>
    public DraftSlide Picture(string name, int x, int y, int width, int height, Stream image)
    {
        this.actions.Add((slide, _) =>
        {
            slide.Shapes.AddPicture(image);

            // Modify the last added picture
            var picture = slide.Shapes[^1];
            picture.Name = name;
            picture.X = x;
            picture.Y = y;
            picture.Width = width;
            picture.Height = height;
        });

        return this;
    }

    /// <summary>
    ///     Adds a picture to the slide, centered on the slide.
    /// </summary>
    public DraftSlide Picture(byte[] imageBytes)
    {
        this.actions.Add((slide, pres) =>
        {
            using var stream = new MemoryStream(imageBytes);
            slide.Shapes.AddPicture(stream);


            var picture = slide.Shapes[^1];
            picture.X = (int)((pres.SlideWidth - picture.Width) / 2);
            picture.Y = (int)((pres.SlideHeight - picture.Height) / 2);
        });

        return this;
    }

    /// <summary>
    ///     Configures a picture using a nested builder.
    /// </summary>
    public DraftSlide Picture(Action<DraftPicture> configure)
    {
        this.actions.Add((slide, _) =>
        {
            var b = new DraftPicture();
            configure(b);
            slide.Shapes.AddPicture(b.ImageStream);
            var pic = slide.Shapes[slide.Shapes.Count - 1];
            pic.Name = b.DraftName;
            pic.X = b.DraftX;
            pic.Y = b.DraftY;
            pic.Width = b.DraftWidth;
            pic.Height = b.DraftHeight;
            if (!string.IsNullOrEmpty(b.GeometryName))
            {
                pic.GeometryType = (Geometry)Enum.Parse(typeof(Geometry), b.GeometryName!.Replace(" ", string.Empty));
            }
        });

        return this;
    }

    /// <summary>
    ///     Adds a text box (auto shape) and sets its content.
    /// </summary>
    /// <param name="content">Text content.</param>
    public DraftSlide TextShape(string content)
    {
        return this.TextShape(content, x: null, y: null, width: 100, height: 50);
    }

    /// <summary>
    ///     Adds a text box (auto shape) at the specified position and sets its content.
    /// </summary>
    /// <param name="content">Text content.</param>
    /// <param name="x">X coordinate in points.</param>
    /// <param name="y">Y coordinate in points.</param>
    public DraftSlide TextShape(string content, int x, int y)
    {
        return this.TextShape(content, x, y, width: 100, height: 50);
    }

    /// <summary>
    ///     Adds a text box (auto shape) at the specified position, with the specified size and content.
    /// </summary>
    /// <param name="content">Text content.</param>
    /// <param name="x">X coordinate in points.</param>
    /// <param name="y">Y coordinate in points.</param>
    /// <param name="width">Width in points.</param>
    /// <param name="height">Height in points.</param>
    public DraftSlide TextShape(string content, int x, int y, int width, int height)
    {
        return this.TextShape(content, x, (int?)y, width, height);
    }

    /// <summary>
    ///     Adds a text box (auto shape) and sets its content.
    /// </summary>
    /// <param name="content">Text content.</param>
    /// <param name="x">X coordinate in points. If <see langword="null"/>, the text box is centered horizontally.</param>
    /// <param name="y">Y coordinate in points. If <see langword="null"/>, the text box is centered vertically.</param>
    /// <param name="width">Width in points.</param>
    /// <param name="height">Height in points.</param>
    public DraftSlide TextShape(string content, int? x, int? y, int width, int height)
    {
        this.actions.Add((slide, pres) =>
        {
            var effectiveX = x ?? (int)((pres.SlideWidth - width) / 2);
            var effectiveY = y ?? (int)((pres.SlideHeight - height) / 2);
            slide.Shapes.AddTextBox(effectiveX, effectiveY, width, height, content);
        });

        return this;
    }

    /// <summary>
    ///     Adds a text box (auto shape) and sets its content.
    /// </summary>
    public DraftSlide TextShape(string name, int x, int y, int width, int height, string content)
    {
        this.actions.Add((slide, _) =>
        {
            slide.Shapes.AddTextBox(x, y, width, height, content);
            var addedShape = slide.Shapes[^1];
            addedShape.Name = name;
        });

        return this;
    }

    /// <summary>
    ///     Configures a text box using a nested builder.
    /// </summary>
    public DraftSlide TextShape(Action<DraftTextBox> configure)
    {
        return this.Shape(t =>
        {
            t.IsTextBox = true;
            configure(t);
        });
    }

    /// <summary>
    ///     Configures a rectangular auto shape and its text box content using a nested builder.
    /// </summary>
    public DraftSlide Shape(Action<DraftTextBox> configure)
    {
        this.actions.Add((slide, _) => AddRectangleShape(slide, configure));

        return this;
    }

    /// <summary>
    ///     Configures a rectangle shape using a nested builder.
    /// </summary>
    public DraftSlide RectangleShape(Action<DraftRectangle> configure)
    {
        this.actions.Add((slide, _) =>
        {
            var builder = new DraftRectangle();
            configure(builder);

            slide.Shapes.AddShape(
                builder.DraftX,
                builder.DraftY,
                builder.DraftWidth,
                builder.DraftHeight,
                Geometry.Rectangle);
            var rectangle = slide.Shapes[^1];
            rectangle.Name = builder.DraftName;

            ApplySolidFill(rectangle, builder.SolidFillDraft);
        });

        return this;
    }

    /// <summary>
    ///     Adds a line shape.
    /// </summary>
    public DraftSlide LineShape(string name, int startPointX, int startPointY, int endPointX, int endPointY)
    {
        this.actions.Add((slide, _) =>
        {
            slide.Shapes.AddLine(startPointX, startPointY, endPointX, endPointY);
            var line = slide.Shapes[^1];
            line.Name = name;
        });

        return this;
    }

    /// <summary>
    ///     Configures a line shape using a nested builder.
    /// </summary>
    public DraftSlide LineShape(Action<DraftLine> configure)
    {
        this.actions.Add((slide, _) => CreateLine(slide, configure));

        return this;
    }

    /// <summary>
    ///     Configures a line shape with an arrow using a nested builder.
    /// </summary>
    public DraftSlide ArrowLineShape(Action<DraftLine> configure)
    {
        this.actions.Add((slide, _) => AddArrowLineShape(slide, configure));

        return this;
    }

    /// <summary>
    ///     Adds a video shape and sets its properties.
    /// </summary>
    /// <param name="name">Requested shape name (ignored to keep a stable "Video" name as used by tests/examples).</param>
    /// <param name="x">X coordinate in points.</param>
    /// <param name="y">Y coordinate in points.</param>
    /// <param name="elementWidth">Width in points.</param>
    /// <param name="elementHeight">Height in points.</param>
    /// <param name="content">Video stream.</param>
    public DraftSlide Video(string name, int x, int y, int elementWidth, int elementHeight, Stream content)
    {
        this.actions.Add((slide, _) =>
        {
            slide.Shapes.AddVideo(x, y, content);
            var media = slide.Shapes[^1];
            media.Name = name;
            media.X = x;
            media.Y = y;
            media.Width = elementWidth;
            media.Height = elementHeight;
        });

        return this;
    }

    /// <summary>
    ///     Adds a table with specified size.
    /// </summary>
    public DraftSlide TableShape(string name, int x, int y, int columnsCount, int rowsCount)
    {
        this.actions.Add((slide, _) =>
        {
            slide.Shapes.AddTable(x, y, columnsCount, rowsCount);
            var table = slide.Shapes[^1];
            table.Name = name;
        });

        return this;
    }

    /// <summary>
    ///     Configures a table using a nested builder.
    /// </summary>
    public DraftSlide TableShape(Action<DraftTable> configure)
    {
        this.actions.Add((slide, _) =>
        {
            var builder = new DraftTable();
            configure(builder);

            var rowsCount = builder.Rows.Count;
            slide.Shapes.AddTable(builder.TableX, builder.TableY, builder.ColumnsCount, rowsCount);
            var tableShape = slide.Shapes[^1];
            ApplyTableCellConfigurations(tableShape.Table!, builder);
        });

        return this;
    }

    /// <summary>
    ///     Configures a table shape using a nested builder.
    /// </summary>
    public DraftSlide TableShape(Action<DraftTableShape> configure)
    {
        this.actions.Add((slide, _) =>
        {
            var shapeBuilder = new DraftTableShape();
            configure(shapeBuilder);

            var tableBuilder = shapeBuilder.DraftTableBuilder;
            if (tableBuilder == null)
            {
                return;
            }

            var rowsCount = tableBuilder.Rows.Count;
            slide.Shapes.AddTable(shapeBuilder.ShapeX, shapeBuilder.ShapeY, tableBuilder.ColumnsCount, rowsCount);
            var tableShape = slide.Shapes[^1];
            ApplyTableCellConfigurations(tableShape.Table!, tableBuilder);
        });

        return this;
    }

    /// <summary>
    ///     Adds a pie chart with specified name.
    /// </summary>
    public DraftSlide PieChartShape(string name)
    {
        this.actions.Add((slide, _) =>
        {
            var categoryValues = new Dictionary<string, double>
            {
                { "Category 1", 40 }, { "Category 2", 30 }, { "Category 3", 30 }
            };
            slide.Shapes.AddPieChart(100, 100, 400, 300, categoryValues, "Series 1", name);
        });

        return this;
    }

    /// <summary>
    ///     Configures a pie chart shape using a nested builder.
    /// </summary>
    /// <param name="configure">An action that configures the pie chart shape via the nested <see cref="DraftPieChartShape"/> builder.</param>
    public DraftSlide PieChartShape(Action<DraftPieChartShape> configure)
    {
        this.actions.Add((slide, _) =>
        {
            var shapeBuilder = new DraftPieChartShape();
            configure(shapeBuilder);

            var chartBuilder = shapeBuilder.DraftPieChartBuilder;
            if (chartBuilder == null)
            {
                return;
            }

            var categoryValues = BuildCategoryValues(chartBuilder);
            var (x, y, width, height) = GetChartDimensions(shapeBuilder, chartBuilder);

            slide.Shapes.AddPieChart(
                x, 
                y, 
                width, 
                height, 
                categoryValues, 
                chartBuilder.SeriesName,
                chartBuilder.ChartName);
        });

        return this;
    }

    /// <summary>
    ///     Configures a bubble chart shape using a nested builder.
    /// </summary>
    /// <param name="configure">An action that configures the bubble chart shape via the nested <see cref="DraftBubbleChartShape"/> builder.</param>
    public DraftSlide BubbleChartShape(Action<DraftBubbleChartShape> configure)
    {
        this.actions.Add((slide, _) =>
        {
            var shapeBuilder = new DraftBubbleChartShape();
            configure(shapeBuilder);

            var chartBuilder = shapeBuilder.DraftBubbleChartBuilder;
            if (chartBuilder == null)
            {
                return;
            }

            slide.Shapes.AddBubbleChart(
                shapeBuilder.ShapeX,
                shapeBuilder.ShapeY,
                shapeBuilder.ShapeWidth,
                shapeBuilder.ShapeHeight,
                chartBuilder.SeriesPoints,
                chartBuilder.SeriesName,
                chartBuilder.ChartName);
        });

        return this;
    }

    /// <summary>
    ///     Adds a clustered bar chart with configuration.
    /// </summary>
    public DraftSlide ClusteredBarChartShape(Action<DraftChart> configure)
    {
        this.actions.Add((slide, _) =>
        {
            var builder = new DraftChart();
            configure(builder);
            slide.Shapes.AddClusteredBarChart(
                builder.ChartX,
                builder.ChartY,
                builder.ChartWidth,
                builder.ChartHeight,
                builder.CategoryNames,
                builder.SeriesDataList,
                builder.ChartName);
        });

        return this;
    }

    /// <summary>
    ///     Configures a clustered bar chart shape using a nested builder.
    /// </summary>
    /// <param name="configure">An action that configures the clustered bar chart shape via the nested <see cref="DraftClusteredBarChartShape"/> builder.</param>
    public DraftSlide ClusteredBarChartShape(Action<DraftClusteredBarChartShape> configure)
    {
        this.actions.Add((slide, _) =>
        {
            var shapeBuilder = new DraftClusteredBarChartShape();
            configure(shapeBuilder);

            var chartBuilder = shapeBuilder.DraftChartBuilder;
            if (chartBuilder == null)
            {
                return;
            }

            var (x, y, width, height) = GetChartDimensions(shapeBuilder, chartBuilder);
            slide.Shapes.AddClusteredBarChart(
                x,
                y,
                width,
                height,
                chartBuilder.CategoryNames,
                chartBuilder.SeriesDataList,
                chartBuilder.ChartName);
        });

        return this;
    }

    /// <summary>
    ///     Sets the slide background to a solid color.
    /// </summary>
    /// <param name="hexColor">Hex color string (e.g., "FF0000" for red).</param>
    public DraftSlide SolidBackground(string hexColor)
    {
        this.actions.Add((slide, _) =>
        {
            slide.Fill.SetColor(hexColor);
        });

        return this;
    }

    /// <summary>
    ///     Sets the slide background to an image.
    /// </summary>
    public DraftSlide ImageBackground(byte[] imageBytes)
    {
        this.actions.Add((slide, _) =>
        {
            slide.Fill.SetPicture(new MemoryStream(imageBytes));
        });

        return this;
    }

    internal void ApplyTo(Presentation presentation)
    {
        // Always add a new slide for each DraftSlide application
        var sdkPres = presentation.PresDocument.PresentationPart!.Presentation!;
        sdkPres.SlideIdList ??= new DocumentFormat.OpenXml.Presentation.SlideIdList();

        var blankLayout = presentation.MasterSlides[0].LayoutSlides.First(l => l.Name == "Blank");
        presentation.Slides.Add(blankLayout.Number);

        // Target the newly added slide
        var slide = presentation.Slides[^1];
        foreach (var action in this.actions)
        {
            action(slide, presentation);
        }
    }
    
    private static Dictionary<string, double> BuildCategoryValues(DraftPieChart chartBuilder)
    {
        var categoryValues = new Dictionary<string, double>();
        var categories = chartBuilder.CategoryNames;
        var values = chartBuilder.SeriesValues;

        var count = Math.Min(categories.Length, values.Length);
        for (var i = 0; i < count; i++)
        {
            categoryValues[categories[i]] = values[i];
        }

        return categoryValues;
    }

    private static (double X, double Y, double Width, double Height) GetChartDimensions(
        DraftPieChartShape shapeBuilder,
        DraftPieChart chartBuilder)
    {
        var defaultChartBuilder = new DraftPieChartShape().DraftPieChartBuilder;

        if (defaultChartBuilder == null)
        {
            return (shapeBuilder.ShapeX, shapeBuilder.ShapeY, shapeBuilder.ShapeWidth, shapeBuilder.ShapeHeight);
        }

        var x = chartBuilder.ChartX == defaultChartBuilder.ChartX ? shapeBuilder.ShapeX : chartBuilder.ChartX;
        var y = chartBuilder.ChartY == defaultChartBuilder.ChartY ? shapeBuilder.ShapeY : chartBuilder.ChartY;
        var width = chartBuilder.ChartWidth == defaultChartBuilder.ChartWidth
            ? shapeBuilder.ShapeWidth
            : chartBuilder.ChartWidth;
        var height = chartBuilder.ChartHeight == defaultChartBuilder.ChartHeight
            ? shapeBuilder.ShapeHeight
            : chartBuilder.ChartHeight;

        return (x, y, width, height);
    }

    private static (int X, int Y, int Width, int Height) GetChartDimensions(
        DraftClusteredBarChartShape shapeBuilder,
        DraftChart chartBuilder)
    {
        var defaultChartBuilder = new DraftChart();

        var x = chartBuilder.ChartX == defaultChartBuilder.ChartX ? shapeBuilder.ShapeX : chartBuilder.ChartX;
        var y = chartBuilder.ChartY == defaultChartBuilder.ChartY ? shapeBuilder.ShapeY : chartBuilder.ChartY;
        var width = chartBuilder.ChartWidth == defaultChartBuilder.ChartWidth
            ? shapeBuilder.ShapeWidth
            : chartBuilder.ChartWidth;
        var height = chartBuilder.ChartHeight == defaultChartBuilder.ChartHeight
            ? shapeBuilder.ShapeHeight
            : chartBuilder.ChartHeight;

        return (x, y, width, height);
    }

    private static void AddRectangleShape(IUserSlide slide, Action<DraftTextBox> configure)
    {
        var builder = new DraftTextBox();
        configure(builder);

        var addedShape = AddRectangleShape(slide, builder);
        addedShape.Name = builder.TextBoxName;

        ApplyDraftFont(addedShape, builder.FontDraft);
        ApplyTextHighlightColor(addedShape, builder.HighlightColor);
        ApplyDraftParagraphs(addedShape, builder.Paragraphs);
        ApplyTextBoxAutofit(addedShape, builder.IsTextBox);
    }

    private static IShape AddRectangleShape(IUserSlide slide, DraftTextBox builder)
    {
        if (builder.IsTextBox)
        {
            slide.Shapes.AddTextBox(
                builder.PosX,
                builder.PosY,
                builder.BoxWidth,
                builder.BoxHeight,
                builder.Content ?? string.Empty);
            return slide.Shapes[^1];
        }

        slide.Shapes.AddShape(builder.PosX, builder.PosY, builder.BoxWidth, builder.BoxHeight, builder.ShapeGeometry);
        var addedShape = slide.Shapes[^1];
        SetTextIfProvided(addedShape, builder.Content);
        return addedShape;
    }

    private static void AddArrowLineShape(IUserSlide slide, Action<DraftLine> configure)
    {
        var (lineShape, draftLine) = CreateLine(slide, configure);

        // Apply default arrow only if the user did not configure any arrow ends
        if (!draftLine.DraftTailEndType.HasValue && !draftLine.DraftHeadEndType.HasValue)
        {
            draftLine.EndArrow(A.LineEndValues.Triangle);
        }

        ApplyArrowEnds(lineShape, draftLine);
    }

    private static (LineShape LineType, DraftLine DraftLine) CreateLine(IUserSlide slide, Action<DraftLine> configure)
    {
        var draftLine = new DraftLine();
        configure(draftLine);

        var startX = draftLine.DraftX;
        var startY = draftLine.DraftY;
        var endX = startX + draftLine.DraftWidth;
        var endY = startY + draftLine.DraftHeight;
        slide.Shapes.AddLine(startX, startY, endX, endY);
        var lineShape = (LineShape)slide.Shapes[^1];
        lineShape.Name = draftLine.DraftName;

        if (draftLine.DraftStroke?.DraftWidthPoints is { } strokeWidthPoints &&
            lineShape.Outline is { } outline)
        {
            outline.Weight = strokeWidthPoints;
        }

        return (lineShape, draftLine);
    }

    private static void ApplyArrowEnds(LineShape lineShape, DraftLine draftLine)
    {
        var pConnectionShape = (DocumentFormat.OpenXml.Presentation.ConnectionShape)lineShape.OpenXmlElement;
        pConnectionShape.ShapeProperties ??= new DocumentFormat.OpenXml.Presentation.ShapeProperties();

        var aOutline = pConnectionShape.ShapeProperties.GetFirstChild<A.Outline>() ??
                       pConnectionShape.ShapeProperties.AppendChild(new A.Outline());

        if (draftLine.DraftTailEndType.HasValue)
        {
            var aTailEnd = aOutline.GetFirstChild<A.TailEnd>() ?? aOutline.AppendChild(new A.TailEnd());
            aTailEnd.Type = draftLine.DraftTailEndType.Value;
        }

        if (draftLine.DraftHeadEndType.HasValue)
        {
            var aHeadEnd = aOutline.GetFirstChild<A.HeadEnd>() ?? aOutline.AppendChild(new A.HeadEnd());
            aHeadEnd.Type = draftLine.DraftHeadEndType.Value;
        }
    }

    private static void ApplySolidFill(IShape shape, DraftSolidFill? draftSolidFill)
    {
        if (draftSolidFill == null)
        {
            return;
        }

        if (!string.IsNullOrEmpty(draftSolidFill.HexColor))
        {
            shape.Fill!.SetColor(draftSolidFill.HexColor!);
        }

        if (draftSolidFill.TransparencyPercent is { } transparencyPercent)
        {
            if (shape is not Shape scShape)
            {
                return;
            }

            SetSolidFillTransparency(scShape.OpenXmlElement, transparencyPercent);
        }
    }

    private static void SetSolidFillTransparency(OpenXmlElement pShapeTreeElement, decimal transparencyPercent)
    {
        if (transparencyPercent is < 0 or > 100)
        {
            throw new ArgumentOutOfRangeException(
                nameof(transparencyPercent),
                transparencyPercent,
                "Transparency percent must be between 0 and 100.");
        }

        var alphaPercent = 100 - transparencyPercent;
        var alphaVal = alphaPercent * 1000;

        var pShapeProperties = pShapeTreeElement.Descendants<DocumentFormat.OpenXml.Presentation.ShapeProperties>()
            .FirstOrDefault();
        var aSolidFill = pShapeProperties?.GetFirstChild<A.SolidFill>();
        if (aSolidFill == null)
        {
            return;
        }

        var aRgb = aSolidFill.GetFirstChild<A.RgbColorModelHex>();
        if (aRgb != null)
        {
            var aAlpha = aRgb.GetFirstChild<A.Alpha>() ?? aRgb.AppendChild(new A.Alpha());
            aAlpha.Val = (int)alphaVal;
            return;
        }

        var aScheme = aSolidFill.GetFirstChild<A.SchemeColor>();
        if (aScheme != null)
        {
            var aAlpha = aScheme.GetFirstChild<A.Alpha>() ?? aScheme.AppendChild(new A.Alpha());
            aAlpha.Val = (int)alphaVal;
        }
    }

    private static void SetTextIfProvided(IShape shape, string? content)
    {
        if (string.IsNullOrEmpty(content))
        {
            return;
        }

        shape.TextBox!.SetText(content!);
    }

    private static void ApplyDraftFont(IShape shape, DraftFont? fontDraft)
    {
        if (fontDraft == null)
        {
            return;
        }

        ApplyDraftFontToParagraph(shape.TextBox!.Paragraphs[0], fontDraft);
    }

    private static void ApplyTextHighlightColor(IShape shape, Color? highlightColor)
    {
        if (!highlightColor.HasValue)
        {
            return;
        }

        shape.TextBox!.Paragraphs[0].Portions[0].TextHighlightColor = highlightColor.Value;
    }

    private static void ApplyDraftParagraphs(IShape shape, IReadOnlyList<DraftParagraph> draftParagraphs)
    {
        if (draftParagraphs.Count == 0)
        {
            return;
        }

        var textBox = shape.TextBox!;
        for (var i = 0; i < draftParagraphs.Count; i++)
        {
            ApplyDraftParagraph(textBox, i, draftParagraphs[i]);
        }
    }

    private static void ApplyDraftParagraph(ITextBox textBox, int paragraphIndex, DraftParagraph draftParagraph)
    {
        if (paragraphIndex > 0)
        {
            textBox.Paragraphs.Add();
        }

        var paragraph = textBox.Paragraphs[paragraphIndex];
        if (!string.IsNullOrEmpty(draftParagraph.Content))
        {
            paragraph.Text = draftParagraph.Content!;
        }

        ApplyDraftFontToParagraph(paragraph, draftParagraph.FontDraft);
        ApplyDraftIndentation(paragraph, draftParagraph.IndentationDraft);

        if (!draftParagraph.IsBulletedList)
        {
            return;
        }

        paragraph.HorizontalAlignment = TextHorizontalAlignment.Left;
        paragraph.Bullet.Type = BulletType.Character;
        paragraph.Bullet.Character = draftParagraph.BulletCharacter;
        paragraph.Bullet.ApplyDefaultSpacing();
    }

    private static void ApplyDraftFontToParagraph(IParagraph paragraph, DraftFont? fontDraft)
    {
        foreach (var font in paragraph.Portions.Where(p => p.Font is not null).Select(p => p.Font))
        {
            // Each draft paragraph should be independent: do not inherit bold from the previous paragraph.
            font!.IsBold = fontDraft?.IsBoldValue ?? false;

            if (fontDraft?.SizeValue is not null)
            {
                font.Size = fontDraft.SizeValue.Value;
            }
        }
    }

    private static void ApplyDraftIndentation(IParagraph paragraph, DraftIndentation? indentationDraft)
    {
        if (indentationDraft == null)
        {
            return;
        }

        if (indentationDraft.BeforeTextPoints.HasValue)
        {
            paragraph.SetLeftMargin(indentationDraft.BeforeTextPoints.Value);
        }
    }

    private static void ApplyTextBoxAutofit(IShape shape, bool isTextBox)
    {
        if (!isTextBox)
        {
            return;
        }

        var scTextBox = (Texts.TextBox)shape.TextBox!;
        scTextBox.DisableWrapping();
        scTextBox.AutofitType = AutofitType.Resize;
    }

    private static void ApplyTableCellConfigurations(ITable table, DraftTable draftTable)
    {
        for (var rowIndex = 0; rowIndex < draftTable.Rows.Count && rowIndex < table.Rows.Count; rowIndex++)
        {
            var draftRow = draftTable.Rows[rowIndex];
            var tableRow = table.Rows[rowIndex];

            for (var cellIndex = 0; cellIndex < draftRow.Cells.Count && cellIndex < tableRow.Cells.Count; cellIndex++)
            {
                var draftCell = draftRow.Cells[cellIndex];
                var tableCell = tableRow.Cells[cellIndex];

                ApplyCellSolidColor(draftCell, tableCell);
                ApplyCellTextContent(draftCell, tableCell);
                ApplyCellFontColor(draftCell, tableCell);
            }
        }
    }

    private static void ApplyCellSolidColor(DraftCell draftCell, ITableCell tableCell)
    {
        if (!string.IsNullOrEmpty(draftCell.SolidColorHex))
        {
            tableCell.Fill.SetColor(draftCell.SolidColorHex!);
        }
    }

    private static void ApplyCellFontColor(DraftCell draftCell, ITableCell tableCell)
    {
        if (!string.IsNullOrEmpty(draftCell.FontColorHex))
        {
            tableCell.TextBox.Paragraphs[0].SetFontColor(draftCell.FontColorHex!);
        }
    }

    private static void ApplyCellTextContent(DraftCell draftCell, ITableCell tableCell)
    {
        if (!string.IsNullOrEmpty(draftCell.TextContent))
        {
            tableCell.TextBox.SetText(draftCell.TextContent!);
        }
    }
}