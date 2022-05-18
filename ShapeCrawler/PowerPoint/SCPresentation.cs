using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shared;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace

namespace ShapeCrawler
{
    /// <inheritdoc cref="IPresentation" />
    [SuppressMessage("ReSharper", "InconsistentNaming", Justification = "SC — ShapeCrawler")]
    public sealed class SCPresentation : IPresentation
    {
        private bool closed;
        private Lazy<Dictionary<int, FontData>> paraLvlToFontData;
        private Lazy<SCSlideSize> slideSize;
        internal ResettableLazy<SlideMasterCollection> SlideMastersValue;

        internal PresentationDocument PresentationDocument { get; private set; }

        internal bool Editable { get; }

        internal List<ChartWorkbook> ChartWorkbooks { get; } = new ();

        internal Dictionary<int, FontData> ParaLvlToFontData => this.paraLvlToFontData.Value;

        #region Public Properties

        public ISlideCollection Slides => new SCSlideCollection(this);

        public int SlideWidth => this.slideSize.Value.Width;

        public int SlideHeight => this.slideSize.Value.Height;

        public ISlideMasterCollection SlideMasters => this.SlideMastersValue.Value;

        public byte[] ByteArray => this.GetByteArray();

        public ISectionCollection Sections => SCSectionCollection.Create(this);

        internal List<ImagePart> ImageParts => this.GetImageParts();

        internal SCSlideCollection SlidesInternal => (SCSlideCollection)this.Slides;

        private byte[] GetByteArray()
        {
            var stream = new MemoryStream();
            this.PresentationDocument.Clone(stream);

            return stream.ToArray();
        }

        #endregion Public Properties

        #region Public Methods

        /// <summary>
        ///     Opens existing presentation from specified file path.
        /// </summary>
        public static IPresentation Open(string pptxPath, in bool isEditable)
        {
            return new SCPresentation(pptxPath, isEditable);
        }

        /// <summary>
        ///     Opens presentation from specified byte array.
        /// </summary>
        public static IPresentation Open(byte[] pptxBytes, in bool isEditable)
        {
            ThrowIfSourceInvalid(pptxBytes);

            var pptxMemoryStream = new MemoryStream();
            pptxMemoryStream.Write(pptxBytes, 0, pptxBytes.Length);

            return Open(pptxMemoryStream, isEditable);
        }

        /// <summary>
        ///     Opens presentation from stream.
        /// </summary>
        public static IPresentation Open(Stream stream, in bool isEditable)
        {
            return new SCPresentation(stream, isEditable);
        }

        /// <inheritdoc/>
        public void Save()
        {
            this.PresentationDocument.Save();
        }

        /// <inheritdoc/>
        public void SaveAs(string filePath)
        {
            this.ChartWorkbooks.ForEach(cw => cw.Close()); // closes cached Excel documents

            PresentationDocument savedDoc = (PresentationDocument)this.PresentationDocument.Clone(filePath);
            this.PresentationDocument.Close();
            this.PresentationDocument = savedDoc;
        }

        /// <inheritdoc/>
        public void SaveAs(Stream stream)
        {
            this.ChartWorkbooks.ForEach(cw => cw.Close());

            PresentationDocument savedDoc = (PresentationDocument)this.PresentationDocument.Clone(stream);
            this.PresentationDocument.Close();
            this.PresentationDocument = savedDoc;
        }

        /// <inheritdoc/>
        public void Close()
        {
            if (this.closed)
            {
                return;
            }

            this.ChartWorkbooks.ForEach(cw => cw.Close());
            this.PresentationDocument.Close();

            this.closed = true;
        }

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

        private SCPresentation(string pptxPath, in bool isEditable)
        {
            this.Editable = isEditable;
            var option = new OpenSettings { AutoSave = false };

            ThrowIfSourceInvalid(pptxPath);

            this.PresentationDocument = PresentationDocument.Open(pptxPath, isEditable, option);
            this.Init();
        }

        private SCPresentation(Stream pptxStream, in bool isEditable)
        {
            this.Editable = isEditable;
            var option = new OpenSettings { AutoSave = false };

            ThrowIfSourceInvalid(pptxStream);

            this.PresentationDocument = PresentationDocument.Open(pptxStream, isEditable, option);
            this.Init();
        }

        private List<ImagePart> GetImageParts()
        {
            IEnumerable<SlidePicture> slidePictures = this.Slides.SelectMany(sp => sp.Shapes).Where(x => x is SlidePicture).OfType<SlidePicture>();

            return slidePictures.Select(x => x.Image.ImagePart).ToList();
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
            if (lvlToFontData.Any(kvp => kvp.Value.FontSize == null))
            {
                A.TextDefault themeTextDefault =
                    pPresentation.PresentationPart.ThemePart.Theme.ObjectDefaults.TextDefault;
                if (themeTextDefault != null)
                {
                    lvlToFontData = FontDataParser.FromCompositeElement(themeTextDefault.ListStyle);
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
            Check.NotNull(stream, nameof(stream));
            ThrowIfPptxSizeLarge(stream.Length);
        }

        private static void ThrowIfSourceInvalid(byte[] bytes)
        {
            Check.NotNull(bytes, nameof(bytes));
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
            var nbSlides = PresentationDocument.PresentationPart.SlideParts.Count();
            if (nbSlides > Limitations.MaxSlidesNumber)
            {
                Close();
                throw SlidesMuchMoreException.FromMax(Limitations.MaxSlidesNumber);
            }
        }

        private void Init()
        {
            this.ThrowIfSlidesNumberLarge();
            this.slideSize = new Lazy<SCSlideSize>(this.GetSlideSize);
            this.SlideMastersValue = new ResettableLazy<SlideMasterCollection>(() => SlideMasterCollection.Create(this));
            this.paraLvlToFontData =
                new Lazy<Dictionary<int, FontData>>(() => ParseFontHeights(this.PresentationDocument.PresentationPart.Presentation));
        }

        private SCSlideSize GetSlideSize()
        {
            var pSlideSize = this.PresentationDocument.PresentationPart!.Presentation.SlideSize;
            var withPx = PixelConverter.HorizontalEmuToPixel(pSlideSize.Cx.Value);
            var heightPx = PixelConverter.VerticalEmuToPixel(pSlideSize.Cy.Value);

            return new SCSlideSize(withPx, heightPx);
        }
    }
}