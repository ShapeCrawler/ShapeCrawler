using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Factories;
using ShapeCrawler.Services;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideMasters;

[SuppressMessage("ReSharper", "InconsistentNaming", Justification = "SC — ShapeCrawler")]
internal class SCSlideMaster : SlideObject, ISlideMaster
{
    private readonly ResettableLazy<List<SCSlideLayout>> slideLayouts;

    internal SCSlideMaster(SCPresentation pres, P.SlideMaster pSlideMaster, int number)
        : base(pres)
    {
        this.Presentation = pres;
        this.PSlideMaster = pSlideMaster;
        this.slideLayouts = new ResettableLazy<List<SCSlideLayout>>(this.GetSlideLayouts);
        this.Number = number;
    }

    public IImage? Background => this.GetBackground();

    public IReadOnlyList<ISlideLayout> SlideLayouts => this.slideLayouts.Value;

    public IShapeCollection Shapes => ShapeCollection.Create(this.PSlideMaster.SlideMasterPart!, this);

    public ITheme Theme => this.GetTheme();

    public override int Number { get; set; }

    internal P.SlideMaster PSlideMaster { get; }

    internal Dictionary<int, FontData> BodyParaLvlToFontData =>
        FontDataParser.FromCompositeElement(this.PSlideMaster.TextStyles!.BodyStyle!);

    internal Dictionary<int, FontData> TitleParaLvlToFontData =>
        FontDataParser.FromCompositeElement(this.PSlideMaster.TextStyles!.TitleStyle!);

    internal ThemePart ThemePart => this.PSlideMaster.SlideMasterPart!.ThemePart!;

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

        // Other
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
    
    private List<SCSlideLayout> GetSlideLayouts()
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