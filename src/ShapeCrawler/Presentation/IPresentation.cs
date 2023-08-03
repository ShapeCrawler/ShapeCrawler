using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Services;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a presentation document.
/// </summary>
public interface IPresentation : IDisposable
{
    /// <summary>
    ///     Gets the presentation slides.
    /// </summary>
    ISlideCollection Slides { get; }

    /// <summary>
    ///     Gets or sets presentation slides width in pixels.
    /// </summary>
    int SlideWidth { get; set; }

    /// <summary>
    ///     Gets or sets the presentation slides height.
    /// </summary>
    int SlideHeight { get; set; }

    /// <summary>
    ///     Gets collection of the slide masters.
    /// </summary>
    ISlideMasterCollection SlideMasters { get; }

    /// <summary>
    ///     Gets a presentation byte array.
    /// </summary>
    byte[] BinaryData { get; }

    /// <summary>
    ///     Gets section collection.
    /// </summary>
    ISectionCollection Sections { get; }

    /// <summary>
    ///     Gets copy of instance of <see cref="DocumentFormat.OpenXml.Packaging.PresentationDocument"/> class.
    /// </summary>
    PresentationDocument SDKPresentationDocument { get; }

    /// <summary>
    ///     Gets Header and Footer manager.
    /// </summary>
    IHeaderAndFooter HeaderAndFooter { get; }

    /// <summary>
    ///     Saves presentation.
    /// </summary>
    void Save();

    /// <summary>
    ///     Saves presentation in specified file path.
    /// </summary>
    void SaveAs(string path);

    /// <summary>
    ///     Saves presentation in specified stream.
    /// </summary>
    void SaveAs(Stream stream);

    /// <summary>
    ///     Closes the presentation.
    /// </summary>
    void Close();
}

/// <inheritdoc cref="IPresentation"/>
public sealed class SCPresentation : IPresentation
{
    private const int MaxSlidesNumber = 300;
    private readonly Lazy<Dictionary<int, FontData>> paraLvlToFontData;
    private readonly Lazy<SCSlideSize> slideSize;
    private readonly ResetableLazy<SCSectionCollection> sectionCollection;
    private readonly ResetableLazy<SCSlideCollection> slideCollection;

    /// <summary>
    ///     Creates a new presentation from specified file path.
    /// </summary>
    public SCPresentation(string path) : this()
    {
        this.SDKPresentationDocumentInternal =
            new Lazy<PresentationDocument>(() => this.CreatePresDocument(path));
    }

    /// <summary>
    ///     Creates a new presentation from specified stream.
    /// </summary>
    public SCPresentation(Stream stream) : this()
    {
        this.SDKPresentationDocumentInternal =
            new Lazy<PresentationDocument>(() => this.CreatePresDocument(stream));
    }

    /// <summary>
    ///     Creates a new presentation.
    /// </summary>
    public SCPresentation()
    {
        this.SDKPresentationDocumentInternal = new Lazy<PresentationDocument>(this.CreateNewPresDocument);

        this.slideSize = new Lazy<SCSlideSize>(this.GetSlideSize);
        this.SlideMasterCollection =
            new ResetableLazy<SCSlideMasterCollection>(() => SCSlideMasterCollection.Create(this));
        this.paraLvlToFontData =
            new Lazy<Dictionary<int, FontData>>(() =>
                ParseFontHeights(this.SDKPresentationDocumentInternal.Value.PresentationPart!.Presentation));
        this.sectionCollection =
            new ResetableLazy<SCSectionCollection>(() => SCSectionCollection.Create(this));
        this.slideCollection = new ResetableLazy<SCSlideCollection>(() => new SCSlideCollection(this));
        this.HeaderAndFooter = new HeaderAndFooter(this);
    }

    private PresentationDocument CreateNewPresDocument()
    {
        var stream = Assembly.GetExecutingAssembly()
            .GetManifestResourceStream("ShapeCrawler.Resources.new-presentation.pptx") !;
        var pptxStream = new MemoryStream();
        stream.CopyTo(pptxStream);

        return PresentationDocument.Open(pptxStream, true);
    }

    private PresentationDocument CreatePresDocument(string pptxPath)
    {
        var pptxStream = File.ReadAllBytes(pptxPath).ToExpandableStream();
        return CreatePresDocument(pptxStream);
    }
    
    private PresentationDocument CreatePresDocument(Stream pptxStream)
    {
        return PresentationDocument.Open(pptxStream, true);
    }

    /// <inheritdoc/>
    public ISlideCollection Slides => this.slideCollection.Value;

    /// <inheritdoc/>
    public int SlideHeight
    {
        get => this.slideSize.Value.Height;
        set => this.SetSlideHeight(value);
    }

    /// <inheritdoc/>
    public int SlideWidth
    {
        get => this.slideSize.Value.Width;
        set => this.SetSlideWidth(value);
    }

    /// <inheritdoc/>
    public ISlideMasterCollection SlideMasters => this.SlideMasterCollection.Value;

    /// <inheritdoc/>
    public byte[] BinaryData => this.GetByteArray();

    /// <inheritdoc/>
    public ISectionCollection Sections => this.sectionCollection.Value;

    /// <inheritdoc/>
    public PresentationDocument SDKPresentationDocument => this.GetSDKPresentation();

    /// <inheritdoc/>
    public IHeaderAndFooter HeaderAndFooter { get; }

    internal ResetableLazy<SCSlideMasterCollection> SlideMasterCollection { get; }

    internal Lazy<PresentationDocument> SDKPresentationDocumentInternal { get; init; }

    internal SCSectionCollection SectionsInternal => (SCSectionCollection)this.Sections;

    internal List<ChartWorkbook> ChartWorkbooks { get; } = new();

    internal Dictionary<int, FontData> ParaLvlToFontData => this.paraLvlToFontData.Value;

    internal List<ImagePart> ImageParts => this.GetImageParts();

    internal SCSlideCollection SlidesInternal => (SCSlideCollection)this.Slides;

    private static int MaxPresentationSize => 250 * 1024 * 1024;

    /// <inheritdoc/>
    public void Save()
    {
        this.ChartWorkbooks.ForEach(chartWorkbook => chartWorkbook.Close());
        this.SDKPresentationDocumentInternal.Value.Save();
    }

    /// <inheritdoc/>
    public void SaveAs(string path)
    {
        this.SDKPresentationDocumentInternal.Value.Clone(path);
    }

    /// <inheritdoc/>
    public void SaveAs(Stream stream)
    {
        this.SDKPresentationDocumentInternal.Value.Clone(stream);
    }

    /// <inheritdoc/>
    public void Close()
    {
        this.ChartWorkbooks.ForEach(cw => cw.Close());
        this.SDKPresentationDocumentInternal.Value.Dispose();
    }

    /// <summary>
    ///     Closes presentation and releases resources.
    /// </summary>
    public void Dispose()
    {
        this.Close();
    }

    private static Dictionary<int, FontData> ParseFontHeights(P.Presentation pPresentation)
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

    private static void ThrowIfSourceInvalid(Stream stream)
    {
        ThrowIfPptxSizeLarge(stream.Length);
    }

    private static void ThrowIfPptxSizeLarge(in long length)
    {
        if (length > MaxPresentationSize)
        {
            throw PresentationIsLargeException.FromMax(MaxPresentationSize);
        }
    }

    private byte[] GetByteArray()
    {
        var stream = new MemoryStream();
        this.ChartWorkbooks.ForEach(c => c.Close());
        this.SDKPresentationDocumentInternal.Value.Clone(stream);

        return stream.ToArray();
    }

    private void SetSlideHeight(int pixel)
    {
        var pSlideSize = this.SDKPresentationDocumentInternal.Value.PresentationPart!.Presentation.SlideSize!;
        var emu = UnitConverter.VerticalPixelToEmu(pixel);

        pSlideSize.Cy = new Int32Value((int)emu);
    }

    private void SetSlideWidth(int pixel)
    {
        var pSlideSize = this.SDKPresentationDocumentInternal.Value.PresentationPart!.Presentation.SlideSize!;
        var emu = UnitConverter.VerticalPixelToEmu(pixel);

        pSlideSize.Cx = new Int32Value((int)emu);
    }

    private PresentationDocument GetSDKPresentation()
    {
        return (PresentationDocument)this.SDKPresentationDocumentInternal.Value.Clone();
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

    private void ThrowIfSlidesNumberLarge()
    {
        var nbSlides = this.SDKPresentationDocumentInternal.Value.PresentationPart!.SlideParts.Count();
        if (nbSlides > MaxSlidesNumber)
        {
            this.Close();
            throw SlidesMuchMoreException.FromMax(MaxSlidesNumber);
        }
    }

    private SCSlideSize GetSlideSize()
    {
        var pSlideSize = this.SDKPresentationDocumentInternal.Value.PresentationPart!.Presentation.SlideSize!;
        var withPx = UnitConverter.HorizontalEmuToPixel(pSlideSize.Cx!.Value);
        var heightPx = UnitConverter.VerticalEmuToPixel(pSlideSize.Cy!.Value);

        return new SCSlideSize(withPx, heightPx);
    }
}