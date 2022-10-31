using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Constants;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

internal abstract class Shape : IShape
{
    protected Shape(OpenXmlCompositeElement pShapeTreeChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject,
        Shape? groupShape)
        : this(pShapeTreeChild, slideObject)
    {
        this.GroupShape = groupShape;
        this.SlideObject = slideObject.Match(slide => slide as SlideObject, layout => layout, master => master);
    }

    protected Shape(OpenXmlCompositeElement pShapeTreeChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject)
    {
        this.PShapeTreesChild = pShapeTreeChild;
        this.SlideObject = slideObject.Match(slide => slide as SlideObject, layout => layout, master => master);
        this.SlideBase = slideObject.Match(slide => slide as SlideObject, layout => layout, master => master);
    }

    public int Id => (int)this.PShapeTreesChild.GetNonVisualDrawingProperties().Id!.Value;

    public string Name => this.PShapeTreesChild.GetNonVisualDrawingProperties().Name!;

    public bool Hidden =>
        this.DefineHidden(); // TODO: the Shape is inherited by LayoutShape, hence do we need this property?

    public string? CustomData
    {
        get => this.GetCustomData();
        set => this.SetCustomData(value ?? throw new ArgumentNullException(nameof(value)));
    }

    public abstract SCShapeType ShapeType { get; }

    public ISlideObject SlideObject { get; }

    public abstract IPlaceholder? Placeholder { get; }

    public virtual SCGeometry GeometryType => this.GetGeometryType();

    public int X
    {
        get => this.GetXCoordinate();
        set => this.SetXCoordinate(value);
    }

    public int Y
    {
        get => this.GetYCoordinate();
        set => this.SetYCoordinate(value);
    }

    public int Height
    {
        get => this.GetHeightPixels();
        set => this.SetHeight(value);
    }

    public int Width
    {
        get => this.GetWidthPixels();
        set => this.SetWidth(value);
    }

    internal SCSlideMaster SlideMasterInternal
    {
        get
        {
            if (this.SlideBase is SCSlide slide)
            {
                return slide.SlideLayoutInternal.SlideMasterInternal;
            }

            if (this.SlideBase is SCSlideLayout layout)
            {
                return layout.SlideMasterInternal;
            }

            var master = (SCSlideMaster)this.SlideBase;
            return master;
        }
    }

    internal OpenXmlCompositeElement PShapeTreesChild { get; }

    internal SlideObject SlideBase { get; }

    internal P.ShapeProperties PShapeProperties => this.PShapeTreesChild.GetFirstChild<P.ShapeProperties>() !;

    private Shape? GroupShape { get; }

    private void SetCustomData(string value)
    {
        string customDataElement =
            $@"<{SCConstants.CustomDataElementName}>{value}</{SCConstants.CustomDataElementName}>";
        this.PShapeTreesChild.InnerXml += customDataElement;
    }

    private string? GetCustomData()
    {
        var pattern = @$"<{SCConstants.CustomDataElementName}>(.*)<\/{SCConstants.CustomDataElementName}>";
        var regex = new Regex(pattern);
        var elementText = regex.Match(this.PShapeTreesChild.InnerXml).Groups[1];
        if (elementText.Value.Length == 0)
        {
            return null;
        }

        return elementText.Value;
    }

    private bool DefineHidden()
    {
        var parsedHiddenValue = this.PShapeTreesChild.GetNonVisualDrawingProperties().Hidden?.Value;
        return parsedHiddenValue is true;
    }

    private void SetXCoordinate(int value)
    {
        if (this.GroupShape is not null)
        {
            throw new RuntimeDefinedPropertyException("X coordinate of grouped shape cannot be changed.");
        }

        var aOffset = this.PShapeTreesChild.Descendants<A.Offset>().FirstOrDefault();
        if (aOffset == null)
        {
            var placeholderShape = ((Placeholder)this.Placeholder!).ReferencedShape;
            placeholderShape.X = value;
        }
        else
        {
            aOffset.X = PixelConverter.HorizontalPixelToEmu(value);
        }
    }

    private int GetXCoordinate()
    {
        var aOffset = this.PShapeTreesChild.Descendants<A.Offset>().FirstOrDefault();
        if (aOffset == null)
        {
            return ((Placeholder)this.Placeholder!).ReferencedShape.X;
        }

        long xEmu = aOffset.X!;

        if (this.GroupShape is not null)
        {
            var aTransformGroup = ((P.GroupShape)this.GroupShape.PShapeTreesChild).GroupShapeProperties!.TransformGroup;
            xEmu = xEmu - aTransformGroup!.ChildOffset!.X! + aTransformGroup!.Offset!.X!;
        }

        return PixelConverter.HorizontalEmuToPixel(xEmu);
    }

    private void SetYCoordinate(long value)
    {
        if (this.GroupShape is not null)
        {
            throw new RuntimeDefinedPropertyException("Y coordinate of grouped shape cannot be changed.");
        }

        var aOffset = this.PShapeTreesChild.Descendants<A.Offset>().First();
        if (this.Placeholder is not null)
        {
            throw new PlaceholderCannotBeChangedException();
        }

        aOffset.Y = PixelConverter.VerticalPixelToEmu(value);
    }

    private int GetYCoordinate()
    {
        var aOffset = this.PShapeTreesChild.Descendants<A.Offset>().FirstOrDefault();
        if (aOffset == null)
        {
            return ((Placeholder)this.Placeholder!).ReferencedShape.Y;
        }

        var yEmu = aOffset.Y!;

        if (this.GroupShape is not null)
        {
            var aTransformGroup =
                ((P.GroupShape)this.GroupShape.PShapeTreesChild).GroupShapeProperties!.TransformGroup!;
            yEmu = yEmu - aTransformGroup.ChildOffset!.Y! + aTransformGroup!.Offset!.Y!;
        }

        return PixelConverter.VerticalEmuToPixel(yEmu);
    }

    private int GetWidthPixels()
    {
        var aExtents = this.PShapeTreesChild.Descendants<A.Extents>().FirstOrDefault();
        if (aExtents == null)
        {
            var placeholder = (Placeholder)this.Placeholder!;
            return placeholder.ReferencedShape.Width;
        }

        return PixelConverter.HorizontalEmuToPixel(aExtents.Cx!);
    }

    private void SetWidth(int pixels)
    {
        var aExtents = this.PShapeTreesChild.Descendants<A.Extents>().FirstOrDefault();
        if (aExtents == null)
        {
            throw new PlaceholderCannotBeChangedException();
        }

        aExtents.Cx = PixelConverter.HorizontalPixelToEmu(pixels);
    }

    private int GetHeightPixels()
    {
        var aExtents = this.PShapeTreesChild.Descendants<A.Extents>().FirstOrDefault();
        if (aExtents == null)
        {
            return ((Placeholder)this.Placeholder!).ReferencedShape.Height;
        }

        return PixelConverter.VerticalEmuToPixel(aExtents!.Cy!);
    }

    private void SetHeight(int pixels)
    {
        var aExtents = this.PShapeTreesChild.Descendants<A.Extents>().FirstOrDefault();
        if (aExtents == null)
        {
            throw new PlaceholderCannotBeChangedException();
        }

        aExtents.Cy = PixelConverter.VerticalPixelToEmu(pixels);
    }

    private SCGeometry GetGeometryType()
    {
        var spPr = this.PShapeTreesChild.Descendants<P.ShapeProperties>().First(); // TODO: optimize
        var aTransform2D = spPr.Transform2D;
        if (aTransform2D != null)
        {
            var aPresetGeometry = spPr.GetFirstChild<A.PresetGeometry>();

            // Placeholder can have transform on the slide, without having geometry
            if (aPresetGeometry == null)
            {
                if (spPr.OfType<A.CustomGeometry>().Any())
                {
                    return SCGeometry.Custom;
                }
            }
            else
            {
                var name = aPresetGeometry.Preset!.Value.ToString();
                Enum.TryParse(name, true, out SCGeometry geometryType);
                return geometryType;
            }
        }

        var placeholder = (Placeholder)this.Placeholder;
        if (placeholder?.ReferencedShape != null)
        {
            return placeholder.ReferencedShape.GeometryType;
        }

        return SCGeometry.Rectangle; // return default
    }
}