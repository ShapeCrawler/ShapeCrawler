using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a Slide Master.
/// </summary>
public interface ISlideMaster
{
    /// <summary>
    ///     Gets background image if slide master has background, otherwise <see langword="null"/>.
    /// </summary>
    IImage? Background { get; }

    /// <summary>
    ///     Gets the collection of Slide Layouts.
    /// </summary>
    IReadOnlyList<ISlideLayout> SlideLayouts { get; }

    /// <summary>
    ///     Gets the collection of master shapes.
    /// </summary>
    IShapeCollection Shapes { get; }

    /// <summary>
    ///     Gets theme.
    /// </summary>
    ITheme Theme { get; }

    /// <summary>
    ///     Gets slide number. Returns <see langword="null"/> if slide master does not have slide number.
    /// </summary>
    IMasterSlideNumber? SlideNumber { get; }

    /// <summary>
    ///     Returns a shape from the slide master.
    /// </summary>
    /// <param name="shape">The name of the shape.</param>
    /// <returns>The requested shape.</returns>
    IShape Shape(string shape);
}

internal sealed class SlideMaster : ISlideMaster
{
    private readonly Lazy<SlideLayouts> layouts;
    private readonly Lazy<MasterSlideNumber?> slideNumber;
    private readonly SlideMasterPart sdkSlideMasterPart;

    internal SlideMaster(SlideMasterPart sdkSlideMasterPart)
    {
        this.sdkSlideMasterPart = sdkSlideMasterPart;
        this.layouts = new Lazy<SlideLayouts>(() => new SlideLayouts(this.sdkSlideMasterPart));
        this.slideNumber = new Lazy<MasterSlideNumber?>(this.CreateSlideNumber);
        this.Shapes = new ShapeCollection(this.sdkSlideMasterPart);
    }

    public IImage? Background => null;
    
    public IReadOnlyList<ISlideLayout> SlideLayouts => this.layouts.Value;
    
    public IShapeCollection Shapes { get; }
    
    public ITheme Theme => new Theme(this.sdkSlideMasterPart, this.sdkSlideMasterPart.ThemePart!.Theme);
    
    public IMasterSlideNumber? SlideNumber => this.slideNumber.Value;
    
    public int Number { get; set; }
    
    public IShape Shape(string shape) => this.Shapes.GetByName(shape);
    
    private MasterSlideNumber? CreateSlideNumber()
    {
        var pSldNum = this.sdkSlideMasterPart.SlideMaster.CommonSlideData!.ShapeTree!
            .Elements<P.Shape>()
            .FirstOrDefault(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == P.PlaceholderValues.SlideNumber);
        
        return pSldNum is null ? null : new MasterSlideNumber(pSldNum);
    }
}