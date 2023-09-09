using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Fonts;
using ShapeCrawler.Services;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

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
    ///     Gets collection of Slide Layouts.
    /// </summary>
    IReadOnlyList<ISlideLayout> SlideLayouts { get; }

    /// <summary>
    ///     Gets collection of master shapes.
    /// </summary>
    IReadOnlyShapes Shapes => new MasterShapes(this);

    /// <summary>
    ///     Gets theme.
    /// </summary>
    ITheme Theme { get; }

    /// <summary>
    ///     Gets slide number. Returns <see langword="null"/> if slide master does not have slide number.
    /// </summary>
    IMasterSlideNumber? SlideNumber { get; }
}

internal sealed class SlideMaster : ISlideMaster
{
    private readonly ResetableLazy<SlideLayouts> layouts;
    private readonly Lazy<MasterSlideNumber?> slideNumber;
    private readonly SlideMasterPart sdkSlideMasterPart;

    internal SlideMaster(SlideMasterPart sdkSlideMasterPart)
    {
        this.sdkSlideMasterPart = sdkSlideMasterPart;
        this.layouts = new ResetableLazy<SlideLayouts>(() => new SlideLayouts(this.sdkSlideMasterPart));
        this.slideNumber = new Lazy<MasterSlideNumber?>(this.CreateSlideNumber);
    }

    public IImage? Background => this.GetBackground();

    public IReadOnlyList<ISlideLayout> SlideLayouts => this.layouts.Value;
    public IReadOnlyShapes Shapes => new MasterShapes(this);

    public ITheme Theme => this.GetTheme();

    public IMasterSlideNumber? SlideNumber => this.slideNumber.Value;

    public int Number { get; set; }

    private SlidePictureImage? GetBackground()
    {
        return null;
    }
    
    private ITheme GetTheme()
    {
        return new SCTheme(this, this.sdkSlideMasterPart.ThemePart!.Theme);
    }
    
    private MasterSlideNumber? CreateSlideNumber()
    {
        var pSldNum = this.sdkSlideMasterPart.SlideMaster.CommonSlideData!.ShapeTree!
            .Elements<P.Shape>()
            .FirstOrDefault(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == P.PlaceholderValues.SlideNumber);
        
        return pSldNum is null ? null : new MasterSlideNumber(pSldNum);
    }
}