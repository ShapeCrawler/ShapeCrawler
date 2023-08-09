using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Services;
using ShapeCrawler.Shapes;
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
    ///     Gets theme.
    /// </summary>
    ITheme Theme { get; }

    /// <summary>
    ///     Gets slide number. Returns <see langword="null"/> if slide master does not have slide number.
    /// </summary>
    IMasterSlideNumber? SlideNumber { get; }
}

internal sealed class SCSlideMaster : ISlideStructure, ISlideMaster
{
    private readonly ResetableLazy<List<SCSlideLayout>> slideLayouts;
    private readonly Lazy<SCMasterSlideNumber?> slideNumber;
    private readonly List<ImagePart> imageParts;
    private readonly TypedOpenXmlPart slideTypedOpenXmlPart;
    private PresentationDocument sdkPresentationDocument;

    internal SCSlideMaster(
        P.SlideMaster pSlideMaster, 
        int number, 
        List<ImagePart> imageParts, 
        TypedOpenXmlPart slideTypedOpenXmlPart, 
        PresentationDocument sdkPresentationDocument)
    {
        this.PSlideMaster = pSlideMaster;
        this.Number = number;
        this.imageParts = imageParts;
        this.slideTypedOpenXmlPart = slideTypedOpenXmlPart;
        this.sdkPresentationDocument = sdkPresentationDocument;
        this.slideLayouts = new ResetableLazy<List<SCSlideLayout>>(this.CreateSlideLayouts);
        this.slideNumber = new Lazy<SCMasterSlideNumber?>(this.CreateSlideNumber);
    }

    public IImage? Background => this.GetBackground();

    public IReadOnlyList<ISlideLayout> SlideLayouts => this.slideLayouts.Value;

    public IShapeCollection Shapes => new SCSlideShapes(this.PSlideMaster.SlideMasterPart!, this, this.slideTypedOpenXmlPart, this.imageParts, this.sdkPresentationDocument);
    
    public IPresentation Presentation { get; } = null!;

    public ITheme Theme => this.GetTheme();

    public IMasterSlideNumber? SlideNumber => this.slideNumber.Value;

    public int Number { get; set; }

    internal Dictionary<int, FontData> BodyParaLvlToFontData =>
        FontDataParser.FromCompositeElement(this.PSlideMaster.TextStyles!.BodyStyle!);

    internal Dictionary<int, FontData> TitleParaLvlToFontData =>
        FontDataParser.FromCompositeElement(this.PSlideMaster.TextStyles!.TitleStyle!);

    internal ThemePart ThemePart => this.PSlideMaster.SlideMasterPart!.ThemePart!;
    
    internal P.SlideMaster PSlideMaster { get; }

    internal SCSlideShapes ShapesInternal => (SCSlideShapes)this.Shapes;
    
    internal TypedOpenXmlPart TypedOpenXmlPart => this.PSlideMaster.SlideMasterPart!;

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
            layouts.Add(new SCSlideLayout(this, layoutPart, number++, this.imageParts, this.sdkPresentationDocument));
        }

        return layouts;
    }
    
    private SCMasterSlideNumber? CreateSlideNumber()
    {
        var pSldNum = this.PSlideMaster.CommonSlideData!.ShapeTree!
            .Elements<P.Shape>()
            .FirstOrDefault(s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == P.PlaceholderValues.SlideNumber);
        
        return pSldNum is null ? null : new SCMasterSlideNumber(pSldNum);
    }
}