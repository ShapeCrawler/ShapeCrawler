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
using ShapeCrawler.Models;
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
    public sealed class SCPresentation : IPresentation // TODO: Make internal
    {
        private bool closed;
        private Lazy<Dictionary<int, FontData>> paraLvlToFontData;
        private Lazy<SlideCollection> slides;
        private Lazy<SlideSizeSc> slideSize;

        internal PresentationDocument presentationDocument;
        internal PresentationPart PresentationPart;

        internal bool Editable { get; }

        internal List<ChartWorkbook> ChartWorkbooks { get; } = new ();

        internal Dictionary<int, FontData> ParaLvlToFontData => paraLvlToFontData.Value;

        #region Public Properties

        public ISlideCollection Slides => slides.Value;

        public int SlideWidth => slideSize.Value.Width;

        public int SlideHeight => slideSize.Value.Height;

        public ISlideMasterCollection SlideMasters => SlideMasterCollection.Create(this);
        
        public List<ImagePart> ImageParts => GetImageParts();

        private List<ImagePart> GetImageParts()
        {
            IEnumerable<SlidePicture> slidePictures = this.Slides.SelectMany(sp => sp.Shapes).Where(x => x is SlidePicture).OfType<SlidePicture>();

            return slidePictures.Select(x => x.Image.ImagePart).ToList();
        }

        #endregion Public Properties

        #region Constructors

        private SCPresentation(string pptxPath, in bool isEditable)
        {
            ThrowIfSourceInvalid(pptxPath);
            presentationDocument = PresentationDocument.Open(pptxPath, isEditable);
            Editable = isEditable;
            Init();
        }

        private SCPresentation(Stream pptxStream, in bool isEditable)
        {
            ThrowIfSourceInvalid(pptxStream);
            Editable = isEditable;
            presentationDocument = PresentationDocument.Open(pptxStream, isEditable);
            Init();
        }

        #endregion Constructors

        #region Public Methods

        public static IPresentation Open(string pptxPath, in bool isEditable)
        {
            return new SCPresentation(pptxPath, isEditable);
        }

        public void Save()
        {
            presentationDocument.Save();
        }

        public void SaveAs(string filePath)
        {
            presentationDocument = (PresentationDocument) presentationDocument.SaveAs(filePath);
        }

        public void SaveAs(Stream stream)
        {
            ChartWorkbooks.ForEach(cw => cw.Close());
            presentationDocument = (PresentationDocument) presentationDocument.Clone(stream);
        }

        public void Close()
        {
            if (this.closed)
            {
                return;
            }

            this.presentationDocument.Close();
            ChartWorkbooks.ForEach(cw => cw.Close());

            this.closed = true;
        }

        public void Dispose()
        {
            this.Close();
        }

        public static SCPresentation Open(byte[] pptxBytes, in bool isEditable)
        {
            ThrowIfSourceInvalid(pptxBytes);

            var pptxMemoryStream = new MemoryStream();
            pptxMemoryStream.Write(pptxBytes, 0, pptxBytes.Length);

            return new SCPresentation(pptxMemoryStream, isEditable);
        }

        public static SCPresentation Open(Stream stream, in bool isEditable)
        {
            return new SCPresentation(stream, isEditable);
        }

        #endregion Public Methods

        #region Private Methods


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

        private SlideCollection GetSlides()
        {
            SlideCollection slideCollection = new(this);

            return slideCollection;
        }
        private void ThrowIfSlidesNumberLarge()
        {
            var nbSlides = presentationDocument.PresentationPart.SlideParts.Count();
            if (nbSlides > Limitations.MaxSlidesNumber)
            {
                Close();
                throw SlidesMuchMoreException.FromMax(Limitations.MaxSlidesNumber);
            }
        }

        private void Init()
        {
            this.ThrowIfSlidesNumberLarge();
            this.slides = new Lazy<SlideCollection>(this.GetSlides);
            this.slideSize = new Lazy<SlideSizeSc>(this.GetSlideSize);
            this.PresentationPart = this.presentationDocument.PresentationPart;
            this.paraLvlToFontData =
                new Lazy<Dictionary<int, FontData>>(() => ParseFontHeights(PresentationPart.Presentation));
        }

        private SlideSizeSc GetSlideSize()
        {
            P.SlideSize pSlideSize = this.presentationDocument.PresentationPart.Presentation.SlideSize;
            return new SlideSizeSc(pSlideSize.Cx.Value, pSlideSize.Cy.Value);
        }

        #endregion Private Methods

        internal void ThrowIfClosed()
        {
            if (this.closed)
            {
                throw new ShapeCrawlerException("The presentation is closed.");
            }
        }
    }
}