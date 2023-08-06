using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Services;
using ShapeCrawler.Shared;

namespace ShapeCrawler;

internal sealed record PresentationCore
{
    private readonly Lazy<Dictionary<int, FontData>> paraLvlToFontData;
    private readonly Lazy<SCSlideSize> slideSize;
    private readonly ResetableLazy<SCSectionCollection> sectionCollection;
    private readonly ResetableLazy<SCSlideCollection> slideCollection;

    internal PresentationCore(byte[] bytes)
        : this(new MemoryStream(bytes))
    {
        this.slideSize = new Lazy<SCSlideSize>(this.GetSlideSize);
        this.SlideMasterCollection =
            new ResetableLazy<SCSlideMasterCollection>(() => SCSlideMasterCollection.Create(this));
        this.paraLvlToFontData =
            new Lazy<Dictionary<int, FontData>>(() =>
                ParseFontHeights(this.SDKPresentation!.PresentationPart!.Presentation));
        this.sectionCollection =
            new ResetableLazy<SCSectionCollection>(() => SCSectionCollection.Create(this));
        this.slideCollection = new ResetableLazy<SCSlideCollection>(() => new SCSlideCollection(this));
        this.HeaderAndFooter = new HeaderAndFooter(this);
    }

    internal PresentationCore(Stream stream)
    {
        this.SDKPresentation = PresentationDocument.Open(stream, true);
        
        this.slideSize = new Lazy<SCSlideSize>(this.GetSlideSize);
        this.SlideMasterCollection =
            new ResetableLazy<SCSlideMasterCollection>(() => SCSlideMasterCollection.Create(this));
        this.paraLvlToFontData =
            new Lazy<Dictionary<int, FontData>>(() =>
                ParseFontHeights(this.SDKPresentation!.PresentationPart!.Presentation));
        this.sectionCollection =
            new ResetableLazy<SCSectionCollection>(() => SCSectionCollection.Create(this));
        this.slideCollection = new ResetableLazy<SCSlideCollection>(() => new SCSlideCollection(this));
        this.HeaderAndFooter = new HeaderAndFooter(this);
    }

    public ISlideCollection Slides => this.slideCollection.Value;

    public int SlideHeight
    {
        get => this.slideSize.Value.Height;
        set => this.SetSlideHeight(value);
    }
    public int SlideWidth
    {
        get => this.slideSize.Value.Width;
        set => this.SetSlideWidth(value);
    }

    public ISlideMasterCollection SlideMasters => this.SlideMasterCollection.Value;

    public byte[] BinaryData => this.GetByteArray();
    
    public ISectionCollection Sections => this.sectionCollection.Value;
    
    public IHeaderAndFooter HeaderAndFooter { get; }

    internal ResetableLazy<SCSlideMasterCollection> SlideMasterCollection { get; }

    internal PresentationDocument SDKPresentation { get; init; }

    internal SCSectionCollection SectionsInternal => (SCSectionCollection)this.Sections;

    internal List<ChartWorkbook> ChartWorkbooks { get; } = new();

    internal Dictionary<int, FontData> ParaLvlToFontData => this.paraLvlToFontData.Value;

    internal List<ImagePart> ImageParts => this.GetImageParts();

    internal SCSlideCollection SlidesInternal => (SCSlideCollection)this.Slides;

    public void Save(string path)
    {
        var cloned = this.SDKPresentation.Clone(path);
        cloned.Dispose();
    }

    public void Save(Stream stream)
    {
        this.SDKPresentation.Clone(stream);
    }

    private static Dictionary<int, FontData> ParseFontHeights(
        DocumentFormat.OpenXml.Presentation.Presentation pPresentation)
    {
        var lvlToFontData = new Dictionary<int, FontData>();

        // from presentation default text settings
        if (pPresentation.DefaultTextStyle != null)
        {
            lvlToFontData = FontDataParser.FromCompositeElement(pPresentation.DefaultTextStyle);
        }

        // from theme default text settings
        if (lvlToFontData.Any(kvp => kvp.Value.FontSize is null))
        {
            var themeTextDefault =
                pPresentation.PresentationPart!.ThemePart!.Theme.ObjectDefaults!.TextDefault;
            if (themeTextDefault != null)
            {
                lvlToFontData = FontDataParser.FromCompositeElement(themeTextDefault.ListStyle!);
            }
        }

        return lvlToFontData;
    }

    private byte[] GetByteArray()
    {
        var stream = new MemoryStream();
        this.ChartWorkbooks.ForEach(c => c.Close());
        this.SDKPresentation.Clone(stream);

        return stream.ToArray();
    }

    private void SetSlideHeight(int pixel)
    {
        var pSlideSize = this.SDKPresentation.PresentationPart!.Presentation.SlideSize!;
        var emu = UnitConverter.VerticalPixelToEmu(pixel);

        pSlideSize.Cy = new Int32Value((int)emu);
    }

    private void SetSlideWidth(int pixel)
    {
        var pSlideSize = this.SDKPresentation.PresentationPart!.Presentation.SlideSize!;
        var emu = UnitConverter.VerticalPixelToEmu(pixel);

        pSlideSize.Cx = new Int32Value((int)emu);
    }

    private List<ImagePart> GetImageParts()
    {
        var allShapes = this.SlidesInternal.SelectMany(slide => slide.Shapes);
        var imgParts = new List<ImagePart>();

        FromShapes(allShapes);

        return imgParts;

        void FromShapes(IEnumerable<IShape> shapes)
        {
            foreach (var shape in shapes)
            {
                switch (shape)
                {
                    case SCPicture slidePicture:
                        imgParts.Add(((SCImage)slidePicture.Image).SDKImagePart);
                        break;
                    case IGroupShape groupShape:
                        FromShapes(groupShape.Shapes.Select(x => x));
                        break;
                }
            }
        }
    }

    private SCSlideSize GetSlideSize()
    {
        var pSlideSize = this.SDKPresentation.PresentationPart!.Presentation.SlideSize!;
        var withPx = UnitConverter.HorizontalEmuToPixel(pSlideSize.Cx!.Value);
        var heightPx = UnitConverter.VerticalEmuToPixel(pSlideSize.Cy!.Value);

        return new SCSlideSize(withPx, heightPx);
    }
}