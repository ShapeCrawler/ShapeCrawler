using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Models.Settings;
using ShapeCrawler.Shared;
using ShapeCrawler.Statics;

namespace ShapeCrawler.Models
{
    /// <summary>
    /// <inheritdoc cref="IPresentation"/>
    /// </summary>
    public class PresentationEx : IPresentation
    {
        #region Fields

        private PresentationDocument _outerSdkPresentation;
        private Lazy<EditAbleCollection<Slide>> _slides;
        private Lazy<SlideSize> _slideSize;
        private bool _closed;
        private PreSettings _preSettings;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets the presentation slides.
        /// </summary>
        public EditAbleCollection<Slide> Slides => _slides.Value;

        /// <summary>
        /// Gets the presentation slides width.
        /// </summary>
        public int SlideWidth => _slideSize.Value.Width;

        /// <summary>
        /// Gets the presentation slides height.
        /// </summary>
        public int SlideHeight => _slideSize.Value.Height;

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationEx"/> class by pptx-file path.
        /// </summary>
        public PresentationEx(string pptxPath)
        {
            ThrowIfSourceInvalid(pptxPath);

            _outerSdkPresentation = PresentationDocument.Open(pptxPath, true);

            Init();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationEx"/> class by pptx-file stream.
        /// </summary>
        public PresentationEx(Stream pptxStream, bool isEditable = false)
        {
            ThrowIfSourceInvalid(pptxStream);
            _outerSdkPresentation = PresentationDocument.Open(pptxStream, isEditable);
            Init();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationEx"/> class by pptx-file stream.
        /// </summary>
        private PresentationEx(MemoryStream pptxStream, bool isEditable)
        {
            ThrowIfSourceInvalid(pptxStream);
            _outerSdkPresentation = PresentationDocument.Open(pptxStream, isEditable);
            Init();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationEx"/> class by pptx-file byte array.
        /// </summary>
        /// <param name="pptxBytes"></param>
        public PresentationEx(byte[] pptxBytes)
        {
            ThrowIfSourceInvalid(pptxBytes);

            var pptxStream = new MemoryStream();
            pptxStream.Write(pptxBytes, 0, pptxBytes.Length);
            _outerSdkPresentation = PresentationDocument.Open(pptxStream, true);
            
            Init();
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Saves presentation in specified file path.
        /// </summary>
        /// <param name="filePath"></param>
        public void SaveAs(string filePath)
        {
            Check.NotEmpty(filePath, nameof(filePath));
            _outerSdkPresentation = (PresentationDocument)_outerSdkPresentation.SaveAs(filePath);
        }

        /// <summary>
        /// Saves presentation in specified stream.
        /// </summary>
        /// <param name="stream"></param>
        public void SaveAs(Stream stream)
        {
            Check.NotNull(stream, nameof(stream));
            _outerSdkPresentation = (PresentationDocument)_outerSdkPresentation.Clone(stream);
        }

        /// <summary>
        /// Closes presentation.
        /// </summary>
        public void Close()
        {
            if (_closed)
            {
                return;
            }

            _outerSdkPresentation.Close();
            if (_preSettings != null)
            {
                foreach (var xlsxDoc in _preSettings.XlsxDocuments.Values)
                {
                    xlsxDoc.Close();
                }
            }

            _closed = true;
        }


        public static PresentationEx Open(byte[] pptxBytes, bool isEditable = false)
        {
            ThrowIfSourceInvalid(pptxBytes);

            var pptxMemoryStream = new MemoryStream();
            pptxMemoryStream.Write(pptxBytes, 0, pptxBytes.Length);

            return new PresentationEx(pptxMemoryStream, isEditable);
        }

        #endregion Public Methods

        #region Private Methods

        private EditAbleCollection<Slide> GetSlides()
        {
            var sdkPrePart = _outerSdkPresentation.PresentationPart;
            _preSettings = new PreSettings(sdkPrePart.Presentation, _slideSize);
            var slideCollection = SlideCollection.Create(sdkPrePart, _preSettings);

            return slideCollection;
        }

        private static void ThrowIfSourceInvalid(string path)
        {
            if (!File.Exists(path))
            {
                throw new FileNotFoundException(nameof(path));
            }
            var  fileInfo = new FileInfo(path);

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

        private static void ThrowIfPptxSizeLarge(long length)
        {
            if (length > Limitations.MaxPresentationSize)
            {
                throw PresentationIsLargeException.FromMax(Limitations.MaxPresentationSize);
            }
        }

        private void ThrowIfSlidesNumberLarge()
        {
            var nbSlides = _outerSdkPresentation.PresentationPart.SlideParts.Count();
            if (nbSlides > Limitations.MaxSlidesNumber)
            {
                Close();
                throw SlidesMuchMoreException.FromMax(Limitations.MaxSlidesNumber);
            }
        }

        private void Init()
        {
            ThrowIfSlidesNumberLarge();
            _slides = new Lazy<EditAbleCollection<Slide>>(GetSlides);
            _slideSize = new Lazy<SlideSize>(ParseSlideSize);
        }

        private SlideSize ParseSlideSize()
        {
            var sdkSldSize = _outerSdkPresentation.PresentationPart.Presentation.SlideSize;
            return new SlideSize(sdkSldSize.Cx.Value, sdkSldSize.Cy.Value);
        }

        #endregion Private Methods
    }
}