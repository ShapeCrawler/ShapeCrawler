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
        private Lazy<SlideCollection> slides;
        private Lazy<SlideSizeSc> slideSize;
        internal ResettableLazy<SlideMasterCollection> slideMasters;

        private SCPresentation(string pptxPath, in bool isEditable)
        {
            ThrowIfSourceInvalid(pptxPath);

            this.PresentationDocument = PresentationDocument.Open(pptxPath, isEditable);
            this.Editable = isEditable;
            this.Init();
        }

        internal PresentationDocument PresentationDocument { get; private set; }

        internal bool Editable { get; }

        internal List<ChartWorkbook> ChartWorkbooks { get; } = new ();

        internal Dictionary<int, FontData> ParaLvlToFontData => this.paraLvlToFontData.Value;

        #region Public Properties

        public ISlideCollection Slides => new SlideCollection(this);

        public int SlideWidth => this.slideSize.Value.Width;

        public int SlideHeight => this.slideSize.Value.Height;

        public ISlideMasterCollection SlideMasters => this.slideMasters.Value;

        public byte[] ByteArray => GetByteArray();

        private byte[] GetByteArray()
        {
            var stream = new MemoryStream();
            this.PresentationDocument.Clone(stream);

            return stream.ToArray();
        }

        #endregion Public Properties

        internal List<ImagePart> ImageParts => this.GetImageParts();

        private SCPresentation(Stream pptxStream, in bool isEditable)
        {
            ThrowIfSourceInvalid(pptxStream);
            this.PresentationDocument = PresentationDocument.Open(pptxStream, isEditable);
            this.Editable = isEditable;
            this.Init();
        }

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
        public static SCPresentation Open(byte[] pptxBytes, in bool isEditable)
        {
            ThrowIfSourceInvalid(pptxBytes);

            var pptxMemoryStream = new MemoryStream();
            pptxMemoryStream.Write(pptxBytes, 0, pptxBytes.Length);

            return new SCPresentation(pptxMemoryStream, isEditable);
        }

        /// <summary>
        ///     Opens presentation from stream.
        /// </summary>
        public static SCPresentation Open(Stream stream, in bool isEditable)
        {
            return new SCPresentation(stream, isEditable);
        }

        public void Save()
        {
            this.PresentationDocument.Save();
        }

        public void SaveAs(string filePath)
        {
            PresentationDocument = (PresentationDocument) PresentationDocument.SaveAs(filePath);
        }

        public void SaveAs(Stream stream)
        {
            this.ChartWorkbooks.ForEach(cw => cw.Close());
            this.PresentationDocument = (PresentationDocument)this.PresentationDocument.Clone(stream);
        }

        public void Close()
        {
            if (this.closed)
            {
                return;
            }

            this.PresentationDocument.Close();
            ChartWorkbooks.ForEach(cw => cw.Close());

            this.closed = true;
        }

        public void Dispose()
        {
            this.Close();
        }

        #endregion Public Methods

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
            this.slideSize = new Lazy<SlideSizeSc>(this.GetSlideSize);
            this.slideMasters = new ResettableLazy<SlideMasterCollection>(() => SlideMasterCollection.Create(this));
            this.paraLvlToFontData =
                new Lazy<Dictionary<int, FontData>>(() => ParseFontHeights(this.PresentationDocument.PresentationPart.Presentation));
        }

        private SlideSizeSc GetSlideSize()
        {
            P.SlideSize pSlideSize = this.PresentationDocument.PresentationPart.Presentation.SlideSize;
            return new SlideSizeSc(pSlideSize.Cx.Value, pSlideSize.Cy.Value);
        }

        internal void ThrowIfClosed()
        {
            if (this.closed)
            {
                throw new ShapeCrawlerException("The presentation is closed.");
            }
        }
    }
}