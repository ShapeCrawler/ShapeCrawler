using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;
using Position = ShapeCrawler.Positions.Position;

namespace ShapeCrawler.ShapeCollection;

internal abstract class Shape : IShape
{
    protected readonly TypedOpenXmlPart sdkTypedOpenXmlPart;
    protected readonly OpenXmlElement pShapeTreeElement;
    private const string customDataElementName = "ctd";
    private readonly Position position;
    private readonly ShapeSize size;
    private readonly ShapeId shapeId;

    internal Shape(TypedOpenXmlPart sdkTypedOpenXmlPart, OpenXmlElement pShapeTreeElement)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.pShapeTreeElement = pShapeTreeElement;
        this.position = new Position(sdkTypedOpenXmlPart, pShapeTreeElement);
        this.size = new ShapeSize(this.sdkTypedOpenXmlPart, pShapeTreeElement);
        this.shapeId = new ShapeId(pShapeTreeElement);
    }

    public virtual int X
    {
        get => this.position.X();
        set => this.position.UpdateX(value);
    }

    public virtual int Y
    {
        get => this.position.Y();
        set => this.position.UpdateY(value);
    }

    public int Width
    {
        get => this.size.Width();
        set => this.size.UpdateWidth(value);
    }

    public int Height
    {
        get => this.size.Height();
        set => this.size.UpdateHeight(value);
    }

    public int Id => this.shapeId.Value();

    public string Name => this.pShapeTreeElement.NonVisualDrawingProperties().Name!.Value!;

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
            var pPlaceholderShape = this.pShapeTreeElement.Descendants<P.PlaceholderShape>().FirstOrDefault();
            if (pPlaceholderShape == null)
            {
                throw new SCException(
                    $"The shape is not a placeholder. Use {nameof(IShape.IsPlaceholder)} property to check if shape is a placeholder.");
            }
            
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
        
            return (PlaceholderType)Enum.Parse(typeof(PlaceholderType), pPlaceholderValue.Value.ToString());
        }
    } 
        

    public virtual Geometry GeometryType => Geometry.Rectangle;

    public string? CustomData
    {
        get
        {
            const string pattern = @$"<{customDataElementName}>(.*)<\/{customDataElementName}>";

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
                $@"<{customDataElementName}>{value}</{customDataElementName}>";
            this.pShapeTreeElement.InnerXml += customDataElement;
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
    
    public virtual ITextFrame TextFrame { get; protected init; } = new NullTextFrame();

    public virtual double Rotation
    {
        get
        {
            var pSpPr = this.pShapeTreeElement.GetFirstChild<P.ShapeProperties>() !;
            var aTransform2D = pSpPr.Transform2D;
            if (aTransform2D == null)
            {
                aTransform2D = new ReferencedPShape(this.sdkTypedOpenXmlPart, this.pShapeTreeElement).ATransform2D();
                var angle2 = aTransform2D.Rotation!.Value; // rotation angle in 1/60,000th of a degree
                return angle2 / 60000d;
            }
            
            var angle = pSpPr.Transform2D!.Rotation!.Value; // rotation angle in 1/60,000th of a degree
            return angle / 60000d;
        }
    }

    public virtual bool Removeable => false;
    
    public string SDKXPath => new XmlPath(this.pShapeTreeElement).XPath;
    
    public virtual ITable AsTable() => throw new SCException(
        $"The shape is not a table. Use {nameof(IShape.ShapeType)} property to check if the shape is a table.");

    public virtual IMediaShape AsMedia() =>
        throw new SCException(
            $"The shape is not a media shape. Use {nameof(IShape.ShapeType)} property to check if the shape is a media (audio, video, etc.");
    
    public virtual void Remove() => this.pShapeTreeElement.Remove();
}