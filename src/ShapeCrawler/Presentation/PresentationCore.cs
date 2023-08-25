using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Drawing;
using ShapeCrawler.Fonts;
using ShapeCrawler.Services;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

internal sealed record PresentationCore
{
    private readonly Lazy<SCSlideSize> slideSize;
    private readonly ResetableLazy<SCSectionCollection> sectionCollection;
    private readonly ResetableLazy<Slides> slides;

    internal PresentationCore(byte[] bytes)
        : this(new MemoryStream(bytes))
    {
    }

    internal PresentationCore(Stream stream)
    {
        this.SDKPresentationDocument = PresentationDocument.Open(stream, true);
        this.slideSize = new Lazy<SCSlideSize>(this.GetSlideSize);
        var sdkMasterParts = this.SDKPresentationDocument.PresentationPart!.SlideMasterParts;
        this.SlideMastersLazy =
            new ResetableLazy<SlideMasterCollection>(() => new SlideMasterCollection(sdkMasterParts, this));
        this.sectionCollection =
            new ResetableLazy<SCSectionCollection>(() => SCSectionCollection.Create(this));
        this.slides = new ResetableLazy<Slides>(() => new Slides(this, this.ImageParts));
        this.HeaderAndFooter = new HeaderAndFooter(this);
    }

    public ISlideCollection Slides => this.slides.Value;

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

    public ISlideMasterCollection SlideMasters => this.SlideMastersLazy.Value;

    public byte[] BinaryData => this.GetByteArray();

    public ISectionCollection Sections => this.sectionCollection.Value;

    public IHeaderAndFooter HeaderAndFooter { get; }

    internal ResetableLazy<SlideMasterCollection> SlideMastersLazy { get; }

    internal PresentationDocument SDKPresentationDocument { get; init; }

    internal SCSectionCollection SectionsInternal => (SCSectionCollection)this.Sections;

    internal List<ChartWorkbook> ChartWorkbooks { get; } = new();

    internal List<ImagePart> ImageParts => this.GetImageParts();

    internal Slides SlidesInternal => (Slides)this.Slides;

    public void Save(string path)
    {
        var cloned = this.SDKPresentationDocument.Clone(path);
        cloned.Dispose();
    }

    public void Save(Stream stream)
    {
        this.SDKPresentationDocument.Clone(stream);
    }

    internal ParagraphLevelFont? FontDataOrNullForParagraphLevel(int paraLevel)
    {
        var pPresentation = this.SDKPresentationDocument.PresentationPart!.Presentation;
        var lvlToFontData = new Dictionary<int, ParagraphLevelFont>();

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

        if (lvlToFontData.ContainsKey(paraLevel))
        {
            return lvlToFontData[paraLevel];
        }

        return null;
    }

    private byte[] GetByteArray()
    {
        var stream = new MemoryStream();
        this.ChartWorkbooks.ForEach(c => c.Close());
        this.SDKPresentationDocument.Clone(stream);

        return stream.ToArray();
    }

    private void SetSlideHeight(int pixel)
    {
        var pSlideSize = this.SDKPresentationDocument.PresentationPart!.Presentation.SlideSize!;
        var emu = UnitConverter.VerticalPixelToEmu(pixel);

        pSlideSize.Cy = new Int32Value((int)emu);
    }

    private void SetSlideWidth(int pixel)
    {
        var pSlideSize = this.SDKPresentationDocument.PresentationPart!.Presentation.SlideSize!;
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
                    case SlidePicture slidePicture:
                        imgParts.Add(((SlidePictureImage)slidePicture.Image).SDKImagePart);
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
        var pSlideSize = this.SDKPresentationDocument.PresentationPart!.Presentation.SlideSize!;
        var withPx = UnitConverter.HorizontalEmuToPixel(pSlideSize.Cx!.Value);
        var heightPx = UnitConverter.VerticalEmuToPixel(pSlideSize.Cy!.Value);

        return new SCSlideSize(withPx, heightPx);
    }

    internal PresentationPart SDKPresentationPart()
    {
        return this.SDKPresentationDocument.PresentationPart!;
    }
}