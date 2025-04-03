using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Slides;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;
using Position = ShapeCrawler.Positions.Position;

namespace ShapeCrawler.Shapes;

internal class Shape : IShape
{
    private readonly OpenXmlElement pShapeTreeElement;
    private readonly Position position;
    private readonly ShapeSize size;
    private readonly ShapeId shapeId;

    internal Shape(OpenXmlElement pShapeTreeElement)
    {
        this.pShapeTreeElement = pShapeTreeElement;
        this.position = new Position(pShapeTreeElement);
        this.size = new ShapeSize(pShapeTreeElement);
        this.shapeId = new ShapeId(pShapeTreeElement);
    }

    internal Shape(P.Shape pShape, TextBox textBox)
        : this(pShape)
    {
        this.pShapeTreeElement = pShape;
        this.position = new Position(pShape);
        this.size = new ShapeSize(pShape);
        this.shapeId = new ShapeId(pShape);
        this.TextBox = textBox;
    }

    public virtual decimal X
    {
        get => this.position.X;
        set => this.position.X = value;
    }

    public virtual decimal Y
    {
        get => this.position.Y;
        set => this.position.Y = value;
    }

    public decimal Width
    {
        get => this.size.Width;
        set => this.size.Width = value;
    }

    public decimal Height
    {
        get => this.size.Height;
        set => this.size.Height = value;
    }

    public IPresentation Presentation =>
        new Presentation(
            (PresentationDocument)this.pShapeTreeElement.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!
                .OpenXmlPackage);

    public int Id
    {
        get => this.shapeId.Value();
        internal set => this.shapeId.Update(value);
    }

    public string Name
    {
        get => this.pShapeTreeElement.NonVisualDrawingProperties().Name!.Value!;
        set => this.pShapeTreeElement.NonVisualDrawingProperties().Name = new StringValue(value);
    }

    public string AltText
    {
        get => this.pShapeTreeElement.NonVisualDrawingProperties().Description?.Value ?? string.Empty;
        set => this.pShapeTreeElement.NonVisualDrawingProperties().Description = new StringValue(value);
    }

    public bool Hidden
    {
        get
        {
            var parsedHiddenValue = this.pShapeTreeElement.NonVisualDrawingProperties().Hidden?.Value;
            return parsedHiddenValue is true;
        }
    }

    public bool IsPlaceholder => this.pShapeTreeElement.Descendants<P.PlaceholderShape>().Any();

    public PlaceholderType PlaceholderType
    {
        get
        {
            var pPlaceholderShape = this.pShapeTreeElement.Descendants<P.PlaceholderShape>().FirstOrDefault() ??
                                    throw new SCException(
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
        get
        {
            var shapeProperties = this.pShapeTreeElement.Descendants<P.ShapeProperties>().First();
            return new ShapeGeometry(shapeProperties).GeometryType;
        }

        set
        {
            var shapeProperties = this.pShapeTreeElement.Descendants<P.ShapeProperties>().First();
            new ShapeGeometry(shapeProperties).GeometryType = value;
        }
    }

    public virtual decimal CornerSize
    {
        get
        {
            var shapeProperties = this.pShapeTreeElement.Descendants<P.ShapeProperties>().First();
            return new ShapeGeometry(shapeProperties).CornerSize;
        }

        set
        {
            var shapeProperties = this.pShapeTreeElement.Descendants<P.ShapeProperties>().First();
            new ShapeGeometry(shapeProperties).CornerSize = value;
        }
    }

    public virtual decimal[] Adjustments
    {
        get
        {
            var shapeProperties = this.pShapeTreeElement.Descendants<P.ShapeProperties>().First();
            return new ShapeGeometry(shapeProperties).Adjustments;
        }

        set
        {
            var shapeProperties = this.pShapeTreeElement.Descendants<P.ShapeProperties>().First();
            new ShapeGeometry(shapeProperties).Adjustments = value;
        }
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

            var elementText = regex.Match(this.pShapeTreeElement.InnerXml).Groups[1];
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
            this.pShapeTreeElement.InnerXml += customDataElement;
        }
    }

    public virtual ShapeContent ShapeContent => ShapeContent.Shape;

    public virtual bool HasOutline => false;

    public virtual IShapeOutline Outline
    {
        get
        {
            var pShapeProperties = this.pShapeTreeElement.GetFirstChild<P.ShapeProperties>() !;
            return new SlideShapeOutline(pShapeProperties);
        }
    }

    public virtual bool HasFill => false;

    public virtual IShapeFill Fill
    {
        get
        {
            var pShapeProperties = this.pShapeTreeElement.GetFirstChild<P.ShapeProperties>() !;
            return new ShapeFill(pShapeProperties);
        }
    }

    public ITextBox? TextBox { get; }

    public virtual double Rotation
    {
        get
        {
            var pSpPr = this.pShapeTreeElement.GetFirstChild<P.ShapeProperties>() !;
            var aTransform2D = pSpPr.Transform2D;
            if (aTransform2D == null)
            {
                aTransform2D = new ReferencedPShape(this.pShapeTreeElement).ATransform2D();
                if (aTransform2D.Rotation is null)
                {
                    return 0;
                }

                return
                    aTransform2D.Rotation.Value /
                    60000d; // OpenXML rotation angles are stored in units of 1/60,000th of a degree
            }

            return pSpPr.Transform2D!.Rotation!.Value / 60000d;
        }
    }

    public virtual bool Removeable => false;

    public string SDKXPath => new XmlPath(this.pShapeTreeElement).XPath;

    public OpenXmlElement SDKOpenXmlElement => this.pShapeTreeElement.CloneNode(true);

    public virtual ITable AsTable() => throw new SCException(
        $"The shape is not a table. Use {nameof(IShape.ShapeContent)} property to check if the shape is a table.");

    public virtual IMediaShape AsMedia() =>
        throw new SCException(
            $"The shape is not a media shape. Use {nameof(IShape.ShapeContent)} property to check if the shape is a media (audio, video, etc.");

    public void Duplicate()
    {
        var pShapeTree = (P.ShapeTree)this.pShapeTreeElement.Parent!;
        new SCPShapeTree(pShapeTree).Add(this.pShapeTreeElement);
    }

    public virtual void Remove() => this.pShapeTreeElement.Remove();

    public void CopyTo(P.ShapeTree pShapeTree) => new SCPShapeTree(pShapeTree).Add(this.pShapeTreeElement);
}