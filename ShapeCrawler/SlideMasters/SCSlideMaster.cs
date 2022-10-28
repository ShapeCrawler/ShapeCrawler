using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Factories;
using ShapeCrawler.Services;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideMasters;

[SuppressMessage("ReSharper", "InconsistentNaming", Justification = "SC — ShapeCrawler")]
internal class SCSlideMaster : SlideObject, ISlideMaster
{
    private readonly ResettableLazy<List<SCSlideLayout>> slideLayouts;

    internal SCSlideMaster(SCPresentation pres, P.SlideMaster pSlideMaster)
        : base(pres)
    {
        this.Presentation = pres;
        this.PSlideMaster = pSlideMaster;
        this.slideLayouts = new ResettableLazy<List<SCSlideLayout>>(this.GetSlideLayouts);
    }

    public IImage? Background => this.GetBackground();

    public IReadOnlyList<ISlideLayout> SlideLayouts => this.slideLayouts.Value;

    public IShapeCollection Shapes => ShapeCollection.Create(this.PSlideMaster.SlideMasterPart!, this);

    public SCPresentation PresentationInternal => this.Presentation; // TODO: make internal

    internal P.SlideMaster PSlideMaster { get; }

    internal SCPresentation Presentation { get; }

    internal Dictionary<int, FontData> BodyParaLvlToFontData =>
        FontDataParser.FromCompositeElement(this.PSlideMaster.TextStyles!.BodyStyle!);

    internal Dictionary<int, FontData> TitleParaLvlToFontData =>
        FontDataParser.FromCompositeElement(this.PSlideMaster.TextStyles!.TitleStyle!);

    internal ThemePart ThemePart => this.PSlideMaster.SlideMasterPart!.ThemePart!;

    internal ShapeCollection ShapesInternal => (ShapeCollection)this.Shapes;

    internal override TypedOpenXmlPart TypedOpenXmlPart => this.PSlideMaster.SlideMasterPart!;

    internal bool TryGetFontSizeFromBody(int paragraphLvl, out int fontSize)
    {
        Dictionary<int, FontData> bodyParaLvlToFontData =
            FontDataParser.FromCompositeElement(this.PSlideMaster.TextStyles!.BodyStyle!);
        if (bodyParaLvlToFontData.TryGetValue(paragraphLvl, out FontData fontData))
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

    private List<SCSlideLayout> GetSlideLayouts()
    {
        var rIdList = this.PSlideMaster.SlideLayoutIdList!.OfType<P.SlideLayoutId>().Select(x => x.RelationshipId!);
        var layouts = new List<SCSlideLayout>(rIdList.Count());
        foreach (var rId in rIdList)
        {
            var layoutPart = (SlideLayoutPart)this.PSlideMaster.SlideMasterPart!.GetPartById(rId.Value!);
            layouts.Add(new SCSlideLayout(this, layoutPart));
        }

        return layouts;
    }
}