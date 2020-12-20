using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Collections;
using SlideDotNet.Exceptions;
using SlideDotNet.Models.Settings;
using SlideDotNet.Shared;
using SlideDotNet.Statics;

namespace SlideDotNet.Models
{
    /// <summary>
    /// <inheritdoc cref="IPresentation"/>
    /// </summary>
    public class PresentationEx : IPresentation
    {
        #region Fields

        private PresentationDocument _sdkPre;
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

            _sdkPre = PresentationDocument.Open(pptxPath, true);

            Init();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationEx"/> class by pptx-file stream.
        /// </summary>
        /// <param name="pptxStream"></param>
        public PresentationEx(Stream pptxStream)
        {
            ThrowIfSourceInvalid(pptxStream);
            
            _sdkPre = PresentationDocument.Open(pptxStream, true);

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
            _sdkPre = PresentationDocument.Open(pptxStream, true);
            
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
            _sdkPre = (PresentationDocument)_sdkPre.SaveAs(filePath);
        }

        /// <summary>
        /// Saves presentation in specified stream.
        /// </summary>
        /// <param name="stream"></param>
        public void SaveAs(Stream stream)
        {
            Check.NotNull(stream, nameof(stream));
            _sdkPre = (PresentationDocument)_sdkPre.Clone(stream);
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

            _sdkPre.Close();
            if (_preSettings != null)
            {
                foreach (var xlsxDoc in _preSettings.XlsxDocuments.Values)
                {
                    xlsxDoc.Close();
                }
            }

            _closed = true;
        }

        #endregion Public Methods

        #region Private Methods

        private EditAbleCollection<Slide> GetSlides()
        {
            var sdkPrePart = _sdkPre.PresentationPart;
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
            var nbSlides = _sdkPre.PresentationPart.SlideParts.Count();
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
            var sdkSldSize = _sdkPre.PresentationPart.Presentation.SlideSize;
            return new SlideSize(sdkSldSize.Cx.Value, sdkSldSize.Cy.Value);
        }

        #endregion Private Methods
    }
}