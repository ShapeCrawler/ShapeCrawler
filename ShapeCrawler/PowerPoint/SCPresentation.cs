using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Collections;
using ShapeCrawler.Constants;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Services;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     <inheritdoc cref="IPresentation"/>
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming", Justification = "SC — ShapeCrawler")]
    public sealed class SCPresentation : IPresentation
    {
        private bool closed;
        private Lazy<Dictionary<int, FontData>> paraLvlToFontData;
        private Lazy<SCSlideSize> slideSize;
        private ResettableLazy<SCSectionCollection> sectionCollectionLazy;
        private ResettableLazy<SCSlideCollection> slideCollectionLazy;
        private Stream? outerStream;
        private string? outerPath;
        private readonly MemoryStream internalStream;

        private SCPresentation(string outerPath)
        {
            this.outerPath = outerPath;

            ThrowIfSourceInvalid(outerPath);
            var pptxBytes = File.ReadAllBytes(outerPath);
            
            this.internalStream = pptxBytes.ToExpandableStream();
            this.SDKPresentation = PresentationDocument.Open(this.internalStream, true);
            this.Init();
        }

        private SCPresentation(Stream sourceStream)
        {
            this.outerStream = sourceStream;
            ThrowIfSourceInvalid(sourceStream);

            this.internalStream = new MemoryStream();
            sourceStream.CopyTo(this.internalStream);
            this.SDKPresentation = PresentationDocument.Open(this.internalStream, true);
            this.Init();
        }

        private SCPresentation(byte[] sourceBytes)
        {
            this.internalStream = sourceBytes.ToExpandableStream();
            this.SDKPresentation = PresentationDocument.Open(this.internalStream, true);
            this.Init();
        }

        /// <inheritdoc/>
        public ISlideCollection Slides => this.slideCollectionLazy.Value;

        /// <inheritdoc/>
        public int SlideWidth => this.slideSize.Value.Width;

        /// <inheritdoc/>
        public int SlideHeight => this.slideSize.Value.Height;

        /// <inheritdoc/>
        public ISlideMasterCollection SlideMasters => this.SlideMastersValue.Value;

        /// <inheritdoc/>
        public byte[] BinaryData => this.GetByteArray();

        /// <inheritdoc/>
        public ISectionCollection Sections => this.sectionCollectionLazy.Value;

        internal ResettableLazy<SlideMasterCollection> SlideMastersValue { get; private set; }
        
        internal PresentationDocument SDKPresentation { get; private set; }

        internal SCSectionCollection SectionsInternal => (SCSectionCollection)this.Sections;

        internal List<ChartWorkbook> ChartWorkbooks { get; } = new ();

        internal Dictionary<int, FontData> ParaLvlToFontData => this.paraLvlToFontData.Value;

        internal List<ImagePart> ImageParts => this.GetImageParts();

        internal SCSlideCollection SlidesInternal => (SCSlideCollection)this.Slides;

        #region Public Methods

        /// <summary>
        ///     Opens existing presentation from specified file path.
        /// </summary>
        public static IPresentation Open(string pptxPath)
        {
            return new SCPresentation(pptxPath);
        }

        /// <summary>
        ///     Opens presentation from specified byte array.
        /// </summary>
        public static IPresentation Open(byte[] pptxBytes)
        {
            ThrowIfSourceInvalid(pptxBytes);

            return new SCPresentation(pptxBytes);
        }

        /// <summary>
        ///     Opens presentation from specified stream.
        /// </summary>
        public static IPresentation Open(Stream pptxStream)
        {
            pptxStream.Position = 0;
            return new SCPresentation(pptxStream);
        }

        /// <inheritdoc/>
        public void Save()
        {
            this.ChartWorkbooks.ForEach(chartWorkbook => chartWorkbook.Close());
            this.SDKPresentation.Save();

            if (this.outerStream != null)
            {
                this.SDKPresentation.Clone(this.outerStream);
            }
            else if (this.outerPath != null)
            {
                var pres = this.SDKPresentation.Clone(this.outerPath);
                pres.Close();
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
            this.SDKPresentation.Close();

            this.closed = true;
        }

        /// <summary>
        ///     Closes presentation and releases resources.
        /// </summary>
        public void Dispose()
        {
            this.Close();
        }

        #endregion Public Methods

        internal void ThrowIfClosed()
        {
            if (this.closed)
            {
                throw new ShapeCrawlerException("The presentation is closed.");
            }
        }

        private byte[] GetByteArray()
        {
            var stream = new MemoryStream();
            this.SDKPresentation.Clone(stream);

            return stream.ToArray();
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
                            imgParts.Add(((SCImage)slidePicture.Image).SDKImagePart);
                            break;
                        case IGroupShape groupShape:
                            FromShapes(groupShape.Shapes.Select(x => x));
                            break;
                    }
                }
            }
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

        private static void ThrowIfSourceInvalid(byte[] bytes)
        {
            ThrowIfPptxSizeLarge(bytes.Length);
        }

        private static void ThrowIfPptxSizeLarge(in long length)
        {
            if (length > Limitations.MaxPresentationSize)
            {
                throw PresentationIsLargeException.FromMax(Limitations.MaxPresentationSize);
            }
        }

        private void ThrowIfSlidesNumberLarge()
        {
            var nbSlides = this.SDKPresentation.PresentationPart!.SlideParts.Count();
            if (nbSlides > Limitations.MaxSlidesNumber)
            {
                this.Close();
                throw SlidesMuchMoreException.FromMax(Limitations.MaxSlidesNumber);
            }
        }

        private void Init()
        {
            this.ThrowIfSlidesNumberLarge();
            this.slideSize = new Lazy<SCSlideSize>(this.GetSlideSize);
            this.SlideMastersValue =
                new ResettableLazy<SlideMasterCollection>(() => SlideMasterCollection.Create(this));
            this.paraLvlToFontData =
                new Lazy<Dictionary<int, FontData>>(() =>
                    ParseFontHeights(this.SDKPresentation.PresentationPart!.Presentation));
            this.sectionCollectionLazy =
                new ResettableLazy<SCSectionCollection>(() => SCSectionCollection.Create(this));
            this.slideCollectionLazy = new ResettableLazy<SCSlideCollection>(() => new SCSlideCollection(this));
        }

        private SCSlideSize GetSlideSize()
        {
            var pSlideSize = this.SDKPresentation.PresentationPart!.Presentation.SlideSize!;
            var withPx = PixelConverter.HorizontalEmuToPixel(pSlideSize.Cx!.Value);
            var heightPx = PixelConverter.VerticalEmuToPixel(pSlideSize.Cy!.Value);

            return new SCSlideSize(withPx, heightPx);
        }
    }
}