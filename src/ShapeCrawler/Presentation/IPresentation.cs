using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Constants;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
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
    private readonly MemoryStream internalStream;
    private readonly Lazy<Dictionary<int, FontData>> paraLvlToFontData;
    private readonly Lazy<SCSlideSize> slideSize;
    private readonly ResetAbleLazy<SCSectionCollection> sectionCollectionLazy;
    private readonly ResetAbleLazy<SCSlideCollection> slideCollectionLazy;
    private bool closed;
    private Stream? outerStream;
    private string? outerPath;

    private SCPresentation(string outerPath)
    {
        this.outerPath = outerPath;

        ThrowIfSourceInvalid(outerPath);
        var pptxBytes = File.ReadAllBytes(outerPath);

        this.internalStream = pptxBytes.ToExpandableStream();
        this.SDKPresentationInternal = PresentationDocument.Open(this.internalStream, true);

        this.ThrowIfSlidesNumberLarge();
        this.slideSize = new Lazy<SCSlideSize>(this.GetSlideSize);
        this.SlideMastersValue =
            new ResetAbleLazy<SCSlideMasterCollection>(() => SCSlideMasterCollection.Create(this));
        this.paraLvlToFontData =
            new Lazy<Dictionary<int, FontData>>(() =>
                ParseFontHeights(this.SDKPresentationInternal.PresentationPart!.Presentation));
        this.sectionCollectionLazy =
            new ResetAbleLazy<SCSectionCollection>(() => SCSectionCollection.Create(this));
        this.slideCollectionLazy = new ResetAbleLazy<SCSlideCollection>(() => new SCSlideCollection(this));
        this.HeaderAndFooter = new HeaderAndFooter(this);
    }

    private SCPresentation(Stream outerStream)
    {
        this.outerStream = outerStream;
        ThrowIfSourceInvalid(outerStream);

        this.internalStream = new MemoryStream();
        outerStream.CopyTo(this.internalStream);
        this.SDKPresentationInternal = PresentationDocument.Open(this.internalStream, true);

        this.ThrowIfSlidesNumberLarge();
        this.slideSize = new Lazy<SCSlideSize>(this.GetSlideSize);
        this.SlideMastersValue =
            new ResetAbleLazy<SCSlideMasterCollection>(() => SCSlideMasterCollection.Create(this));
        this.paraLvlToFontData =
            new Lazy<Dictionary<int, FontData>>(() =>
                ParseFontHeights(this.SDKPresentationInternal.PresentationPart!.Presentation));
        this.sectionCollectionLazy =
            new ResetAbleLazy<SCSectionCollection>(() => SCSectionCollection.Create(this));
        this.slideCollectionLazy = new ResetAbleLazy<SCSlideCollection>(() => new SCSlideCollection(this));
        this.HeaderAndFooter = new HeaderAndFooter(this);
    }

    /// <inheritdoc/>
    public ISlideCollection Slides => this.slideCollectionLazy.Value;

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
    public ISlideMasterCollection SlideMasters => this.SlideMastersValue.Value;

    /// <inheritdoc/>
    public byte[] BinaryData => this.GetByteArray();

    /// <inheritdoc/>
    public ISectionCollection Sections => this.sectionCollectionLazy.Value;

    /// <inheritdoc/>
    public PresentationDocument SDKPresentationDocument => this.GetSDKPresentation();

    /// <inheritdoc/>
    public IHeaderAndFooter HeaderAndFooter { get; }

    internal ResetAbleLazy<SCSlideMasterCollection> SlideMastersValue { get; }

    internal PresentationDocument SDKPresentationInternal { get; }

    internal SCSectionCollection SectionsInternal => (SCSectionCollection)this.Sections;

    internal List<ChartWorkbook> ChartWorkbooks { get; } = new();

    internal Dictionary<int, FontData> ParaLvlToFontData => this.paraLvlToFontData.Value;

    internal List<ImagePart> ImageParts => this.GetImageParts();

    internal SCSlideCollection SlidesInternal => (SCSlideCollection)this.Slides;

    #region Global Statics

    /// <summary>
    ///     Creates a new presentation.
    /// </summary>
    public static IPresentation Create()
    {
        SCLogger.Send();

        var assembly = Assembly.GetExecutingAssembly();
        var rStream = assembly.GetManifestResourceStream("ShapeCrawler.Resources.new-presentation.pptx") !;
        var mStream = new MemoryStream();
        rStream.CopyTo(mStream);

        return Open(mStream);
    }

    /// <summary>
    ///     Opens presentation path.
    /// </summary>
    public static IPresentation Open(string pptxPath)
    {
        SCLogger.Send();

        return new SCPresentation(pptxPath);
    }

    /// <summary>
    ///     Opens presentation stream.
    /// </summary>
    public static IPresentation Open(Stream pptxStream)
    {
        SCLogger.Send();

        pptxStream.Position = 0;
        return new SCPresentation(pptxStream);
    }

    #endregion Global Statics

    /// <inheritdoc/>
    public void Save()
    {
        this.ChartWorkbooks.ForEach(chartWorkbook => chartWorkbook.Close());
        this.SDKPresentationInternal.Save();

        if (this.outerStream != null)
        {
            this.SDKPresentationInternal.Clone(this.outerStream);
        }
        else if (this.outerPath != null)
        {
            var pres = this.SDKPresentationInternal.Clone(this.outerPath);
            pres.Dispose();
        }
    }

    /// <inheritdoc/>
    public void SaveAs(string path)
    {
        this.outerStream = null;
        this.outerPath = path;
        this.Save();
    }

    /// <inheritdoc/>
    public void SaveAs(Stream stream)
    {
        this.outerPath = null;
        this.outerStream = stream;
        this.Save();
    }

    /// <inheritdoc/>
    public void Close()
    {
        if (this.closed)
        {
            return;
        }

        this.ChartWorkbooks.ForEach(cw => cw.Close());
        this.SDKPresentationInternal.Dispose();

        this.closed = true;
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

    private static void ThrowIfSourceInvalid(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException(nameof(path));
        }

        var fileInfo = new FileInfo(path);

        ThrowIfPptxSizeLarge(fileInfo.Length);
    }

    private static void ThrowIfSourceInvalid(Stream stream)
    {
        ThrowIfPptxSizeLarge(stream.Length);
    }

    private static void ThrowIfPptxSizeLarge(in long length)
    {
        if (length > Limitations.MaxPresentationSize)
        {
            throw PresentationIsLargeException.FromMax(Limitations.MaxPresentationSize);
        }
    }

    private byte[] GetByteArray()
    {
        var stream = new MemoryStream();
        this.ChartWorkbooks.ForEach(c => c.Close());
        this.SDKPresentationInternal.Clone(stream);

        return stream.ToArray();
    }

    private void SetSlideHeight(int pixel)
    {
        var pSlideSize = this.SDKPresentationInternal.PresentationPart!.Presentation.SlideSize!;
        var emu = UnitConverter.VerticalPixelToEmu(pixel);

        pSlideSize.Cy = new Int32Value((int)emu);
    }

    private void SetSlideWidth(int pixel)
    {
        var pSlideSize = this.SDKPresentationInternal.PresentationPart!.Presentation.SlideSize!;
        var emu = UnitConverter.VerticalPixelToEmu(pixel);

        pSlideSize.Cx = new Int32Value((int)emu);
    }
    
    private PresentationDocument GetSDKPresentation()
    {
        return (PresentationDocument)this.SDKPresentationInternal.Clone();
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
        var nbSlides = this.SDKPresentationInternal.PresentationPart!.SlideParts.Count();
        if (nbSlides > Limitations.MaxSlidesNumber)
        {
            this.Close();
            throw SlidesMuchMoreException.FromMax(Limitations.MaxSlidesNumber);
        }
    }

    private SCSlideSize GetSlideSize()
    {
        var pSlideSize = this.SDKPresentationInternal.PresentationPart!.Presentation.SlideSize!;
        var withPx = UnitConverter.HorizontalEmuToPixel(pSlideSize.Cx!.Value);
        var heightPx = UnitConverter.VerticalEmuToPixel(pSlideSize.Cy!.Value);

        return new SCSlideSize(withPx, heightPx);
    }
}