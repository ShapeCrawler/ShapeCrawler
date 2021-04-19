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
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public sealed class SCPresentation : IPresentation
    {
        private bool _closed;

        private Lazy<Dictionary<int, FontData>> _paraLvlToFontData;
        private PresentationDocument _presentationDocument;
        private Lazy<SlideCollection> _slides;
        private Lazy<SlideSizeSc> _slideSize;
        internal PresentationPart PresentationPart;
        internal bool Editable { get; }
        internal List<ChartWorkbook> ChartWorkbooks { get; } = new();

        internal Dictionary<int, FontData> ParaLvlToFontData => _paraLvlToFontData.Value;

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

        #region Public Properties

        public ISlideCollection Slides => _slides.Value;

        public int SlideWidth => _slideSize.Value.Width;

        public int SlideHeight => _slideSize.Value.Height;

        public SlideMasterCollection SlideMasters => SlideMasterCollection.Create(this);

        #endregion Public Properties

        #region Constructors

        private SCPresentation(string pptxPath, in bool isEditable)
        {
            ThrowIfSourceInvalid(pptxPath);
            _presentationDocument = PresentationDocument.Open(pptxPath, isEditable);
            Editable = isEditable;
            Init();
        }

        private SCPresentation(Stream pptxStream, in bool isEditable)
        {
            ThrowIfSourceInvalid(pptxStream);
            Editable = isEditable;
            _presentationDocument = PresentationDocument.Open(pptxStream, isEditable);
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
            _presentationDocument.Save();
        }

        public void SaveAs(string filePath)
        {
            _presentationDocument = (PresentationDocument) _presentationDocument.SaveAs(filePath);
        }

        public void SaveAs(Stream stream)
        {
            ChartWorkbooks.ForEach(cw => cw.Close());
            _presentationDocument = (PresentationDocument) _presentationDocument.Clone(stream);
        }

        public void Close()
        {
            if (_closed)
            {
                return;
            }

            _presentationDocument.Close();
            ChartWorkbooks.ForEach(cw => cw.Close());

            _closed = true;
        }

        public void Dispose()
        {
            Close();
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

        private SlideCollection GetSlides()
        {
            SlideCollection slideCollection = new(this);

            return slideCollection;
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
            var nbSlides = _presentationDocument.PresentationPart.SlideParts.Count();
            if (nbSlides > Limitations.MaxSlidesNumber)
            {
                Close();
                throw SlidesMuchMoreException.FromMax(Limitations.MaxSlidesNumber);
            }
        }

        private void Init()
        {
            ThrowIfSlidesNumberLarge();
            _slides = new Lazy<SlideCollection>(GetSlides);
            _slideSize = new Lazy<SlideSizeSc>(GetSlideSize);
            PresentationPart = _presentationDocument.PresentationPart;
            _paraLvlToFontData =
                new Lazy<Dictionary<int, FontData>>(() => ParseFontHeights(PresentationPart.Presentation));
        }

        private SlideSizeSc GetSlideSize()
        {
            P.SlideSize pSlideSize = _presentationDocument.PresentationPart.Presentation.SlideSize;
            return new SlideSizeSc(pSlideSize.Cx.Value, pSlideSize.Cy.Value);
        }

        #endregion Private Methods
    }
}