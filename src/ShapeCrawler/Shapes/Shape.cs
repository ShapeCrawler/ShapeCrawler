using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;
using Position = ShapeCrawler.Positions.Position;

namespace ShapeCrawler.Shapes;

internal abstract class Shape : IShape
{
    protected readonly OpenXmlPart OpenXmlPart;
    protected readonly OpenXmlElement PShapeTreeElement;
    private readonly Position position;
    private readonly ShapeSize size;
    private readonly ShapeId shapeId;

    internal Shape(OpenXmlPart openXmlPart, OpenXmlElement pShapeTreeElement)
    {
        this.OpenXmlPart = openXmlPart;
        this.PShapeTreeElement = pShapeTreeElement;
        this.position = new Position(openXmlPart, pShapeTreeElement);
        this.size = new ShapeSize(this.OpenXmlPart, pShapeTreeElement);
        this.shapeId = new ShapeId(pShapeTreeElement);
    }

    public virtual decimal X
    {
        get => this.position.X();
        set => this.position.UpdateX(value);
    }

    public virtual decimal Y
    {
        get => this.position.Y();
        set => this.position.UpdateY(value);
    }

    public decimal Width
    {
        get => this.size.Width();
        set => this.size.UpdateWidth(value);
    }

    public decimal Height
    {
        get => this.size.Height();
        set => this.size.UpdateHeight(value);
    }

    public int Id => this.shapeId.Value();

    public string Name
    {
        get => this.PShapeTreeElement.NonVisualDrawingProperties().Name!.Value!;
        set => this.PShapeTreeElement.NonVisualDrawingProperties().Name = new StringValue(value);
    }

    public string AltText
    {
        get => this.PShapeTreeElement.NonVisualDrawingProperties().Description?.Value ?? string.Empty;
        set => this.PShapeTreeElement.NonVisualDrawingProperties().Description = new StringValue(value);
    }

    public bool Hidden
    {
        get
        {
            var parsedHiddenValue = this.PShapeTreeElement.NonVisualDrawingProperties().Hidden?.Value;
            return parsedHiddenValue is true;
        }
    }

    public bool IsPlaceholder => this.PShapeTreeElement.Descendants<P.PlaceholderShape>().Any();

    public PlaceholderType PlaceholderType
    {
        get
        {
            var pPlaceholderShape = this.PShapeTreeElement.Descendants<P.PlaceholderShape>().FirstOrDefault() ?? throw new SCException(
                    $"The shape is not a placeholder. Use {nameof(IShape.IsPlaceholder)} property to check if shape is a placeholder.");
            var pPlaceholderValue = pPlaceholderShape.Type;
            if (pPlaceholderValue == null)
            {
                return PlaceholderType.Content;
            }

            if (pPlaceholderValue == P.PlaceholderValues.Title)
            {
                return PlaceholderType.Title;
            }

            if (pPlaceholderValue == P.PlaceholderValues.CenteredTitle)
            {
                return PlaceholderType.CenteredTitle;
            }

            if (pPlaceholderValue == P.PlaceholderValues.Body)
            {
                return PlaceholderType.Text;
            }

            if (pPlaceholderValue == P.PlaceholderValues.Diagram)
            {
                return PlaceholderType.SmartArt;
            }

            if (pPlaceholderValue == P.PlaceholderValues.ClipArt)
            {
                return PlaceholderType.OnlineImage;
            }

            var value = pPlaceholderValue.ToString() !;

            if (value == "dt")
            {
                return PlaceholderType.DateAndTime;
            }

            if (value == "ftr")
            {
                return PlaceholderType.Footer;
            }

            if (value == "sldNum")
            {
                return PlaceholderType.SlideNumber;
            }

            if (value == "pic")
            {
                return PlaceholderType.Picture;
            }

            if (value == "tbl")
            {
                return PlaceholderType.Table;
            }

            return (PlaceholderType)Enum.Parse(typeof(PlaceholderType), value, true);
        }
    }

    public virtual Geometry GeometryType
    {
        get => Geometry.Rectangle;
        set => throw new SCException("Changing geometry of this shape is not supported");
    }

    public virtual decimal CornerSize
    {
        get => 0;
        set => throw new SCException("Changing corner size of this shape is not supported");
    }

    public virtual decimal[] Adjustments
    {
        get => [];
        set => throw new SCException("Setting adjustments of this shape is not supported");
    }

    public string? CustomData
    {
        get
        {
            const string pattern = @$"<{"ctd"}>(.*)<\/{"ctd"}>";

#if NETSTANDARD2_0
            var regex = new Regex(pattern, RegexOptions.None, TimeSpan.FromSeconds(100));
#else
            var regex = new Regex(pattern, RegexOptions.NonBacktracking);
#endif

            var elementText = regex.Match(this.PShapeTreeElement.InnerXml).Groups[1];
            if (elementText.Value.Length == 0)
            {
                return null;
            }

            return elementText.Value;
        }

        set
        {
            var customDataElement =
                $@"<{"ctd"}>{value}</{"ctd"}>";
            this.PShapeTreeElement.InnerXml += customDataElement;
        }
    }

    public abstract ShapeType ShapeType { get; }

    public virtual bool HasOutline => false;

    public virtual IShapeOutline Outline => throw new SCException(
        $"Shape does not have outline. Use {nameof(IShape.HasOutline)} property to check if the shape has outline.");

    public virtual bool HasFill => false;

    public virtual IShapeFill Fill =>
        throw new SCException(
            $"Shape does not have fill. Use {nameof(IShape.HasFill)} property to check if the shape has fill.");

    public virtual bool IsTextHolder { get; protected init; }

    public virtual ITextBox TextBox { get; protected init; } = default(NullTextFrame);

    public virtual double Rotation
    {
        get
        {
            var pSpPr = this.PShapeTreeElement.GetFirstChild<P.ShapeProperties>() !;
            var aTransform2D = pSpPr.Transform2D;
            if (aTransform2D == null)
            {
                aTransform2D = new ReferencedPShape(this.OpenXmlPart, this.PShapeTreeElement).ATransform2D();
                if (aTransform2D.Rotation is null)
                {
                    return 0;
                }

                return aTransform2D.Rotation.Value / 60000d; // OpenXML rotation angles are stored in units of 1/60,000th of a degree
            }

            return pSpPr.Transform2D!.Rotation!.Value / 60000d;
        }
    }

    public virtual bool Removeable => false;

    public string SdkXPath => new XmlPath(this.PShapeTreeElement).XPath;

    public OpenXmlElement SdkOpenXmlElement => this.PShapeTreeElement.CloneNode(true);

    public string Text
    {
        get => this.TextBox.Text;
        set => this.TextBox.Text = value;
    }

    public virtual ITable AsTable() => throw new SCException(
        $"The shape is not a table. Use {nameof(IShape.ShapeType)} property to check if the shape is a table.");

    public virtual IMediaShape AsMedia() =>
        throw new SCException(
            $"The shape is not a media shape. Use {nameof(IShape.ShapeType)} property to check if the shape is a media (audio, video, etc.");

    public virtual void Remove() => this.PShapeTreeElement.Remove();
}