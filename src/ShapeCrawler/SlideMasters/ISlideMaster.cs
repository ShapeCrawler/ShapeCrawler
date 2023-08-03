using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Services;
using ShapeCrawler.Services.Factories;
using ShapeCrawler.Shared;
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
    ///     Gets collection of shape.
    /// </summary>
    IShapeCollection Shapes { get; }

    /// <summary>
    ///     Gets parent Presentation.
    /// </summary>
    IPresentation Presentation { get; }

    /// <summary>
    ///     Gets theme.
    /// </summary>
    ITheme Theme { get; }

    /// <summary>
    ///     Gets slide number. Returns <see langword="null"/> if slide master does not have slide number.
    /// </summary>
    IMasterSlideNumber? SlideNumber { get; }
}

internal sealed class SCSlideMaster : SlideStructure, ISlideMaster
{
    private readonly ResetableLazy<List<SCSlideLayout>> slideLayouts;
    private readonly Lazy<SCMasterSlideNumber?> slideNumber;

    internal SCSlideMaster(SCPresentation pres, P.SlideMaster pSlideMaster, int number)
        : base(pres)
    {
        this.Presentation = pres;
        this.PSlideMaster = pSlideMaster;
        this.Number = number;
        this.slideLayouts = new ResetableLazy<List<SCSlideLayout>>(this.CreateSlideLayouts);
        this.slideNumber = new Lazy<SCMasterSlideNumber?>(this.CreateSlideNumber);
    }

    private SCMasterSlideNumber? CreateSlideNumber()
    {
        var pSldNum = PSlideMaster.CommonSlideData!.ShapeTree!
            .Elements<P.Shape>()
            .FirstOrDefault(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == P.PlaceholderValues.SlideNumber);
        
        return pSldNum is null ? null : new SCMasterSlideNumber(pSldNum);
    }

    public IImage? Background => this.GetBackground();

    public IReadOnlyList<ISlideLayout> SlideLayouts => this.slideLayouts.Value;

    public override IShapeCollection Shapes => new ShapeCollection(this.PSlideMaster.SlideMasterPart!, this);

    public ITheme Theme => this.GetTheme();

    public IMasterSlideNumber? SlideNumber => this.slideNumber.Value;

    public override int Number { get; set; }

    internal Dictionary<int, FontData> BodyParaLvlToFontData =>
        FontDataParser.FromCompositeElement(this.PSlideMaster.TextStyles!.BodyStyle!);

    internal Dictionary<int, FontData> TitleParaLvlToFontData =>
        FontDataParser.FromCompositeElement(this.PSlideMaster.TextStyles!.TitleStyle!);

    internal ThemePart ThemePart => this.PSlideMaster.SlideMasterPart!.ThemePart!;
    
    internal P.SlideMaster PSlideMaster { get; }

    internal ShapeCollection ShapesInternal => (ShapeCollection)this.Shapes;
    
    internal override TypedOpenXmlPart TypedOpenXmlPart => this.PSlideMaster.SlideMasterPart!;

    internal bool TryGetFontSizeFromBody(int paragraphLvl, out int fontSize)
    {
        var bodyParaLvlToFontData =
            FontDataParser.FromCompositeElement(this.PSlideMaster.TextStyles!.BodyStyle!);
        if (bodyParaLvlToFontData.TryGetValue(paragraphLvl, out var fontData))
        {
            if (fontData.FontSize is not null)
            {
                fontSize = fontData.FontSize;
                return true;
            }
        }

        fontSize = -1;
        return false;
    }

    internal bool TryGetFontSizeFromOther(int paragraphLvl, out int fontSize)
    {
        var pTextStyles = this.PSlideMaster.TextStyles!;

        var otherStyleLvlToFontData = FontDataParser.FromCompositeElement(pTextStyles.OtherStyle!);
        if (otherStyleLvlToFontData.ContainsKey(paragraphLvl))
        {
            if (otherStyleLvlToFontData[paragraphLvl].FontSize is not null)
            {
                fontSize = otherStyleLvlToFontData[paragraphLvl].FontSize!;
                return true;
            }
        }

        fontSize = -1;
        return false;
    }

    private SCImage? GetBackground()
    {
        return null;
    }
    
    private ITheme GetTheme()
    {
        return new SCTheme(this, this.PSlideMaster.SlideMasterPart!.ThemePart!.Theme);
    }
    
    private List<SCSlideLayout> CreateSlideLayouts()
    {
        var rIdList = this.PSlideMaster.SlideLayoutIdList!.OfType<P.SlideLayoutId>().Select(layoutId => layoutId.RelationshipId!);
        var layouts = new List<SCSlideLayout>(rIdList.Count());
        var number = 1;
        foreach (var rId in rIdList)
        {
            var layoutPart = (SlideLayoutPart)this.PSlideMaster.SlideMasterPart!.GetPartById(rId.Value!);
            layouts.Add(new SCSlideLayout(this, layoutPart, number++));
        }

        return layouts;
    }
}