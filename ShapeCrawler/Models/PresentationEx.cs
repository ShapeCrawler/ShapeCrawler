using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using ShapeCrawler.Statics;

namespace ShapeCrawler.Models
{
    public class PresentationEx
    {
        #region Fields
// TODO: Implement IDisposable
        private PresentationDocument _outerSdkPresentation;
        private Lazy<EditableCollection<SlideEx>> _slides;
        private Lazy<SlideSize> _slideSize;
        private bool _closed;
        private PresentationData _preData;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets the presentation slides.
        /// </summary>
        public EditableCollection<SlideEx> Slides => _slides.Value;

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
        internal PresentationEx(string pptxPath, bool isEditable)
        {
            ThrowIfSourceInvalid(pptxPath);
            _outerSdkPresentation = PresentationDocument.Open(pptxPath, isEditable);
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

        public static PresentationEx Open(string pptxPath, bool isEditable)
        {
            return new PresentationEx(pptxPath, isEditable);
        }

        public void Save()
        {
            _outerSdkPresentation.Save();
        }

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
        /// Saves and closes the presentation.
        /// </summary>
        public void Close()
        {
            if (_closed)
            {
                return;
            }

            _outerSdkPresentation.Close();
            if (_preData != null)
            {
                foreach (var xlsxDoc in _preData.XlsxDocuments.Values)
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

        public static PresentationEx Open(Stream stream, bool isEditable)
        {
            return new PresentationEx(stream, isEditable);
        }

        #endregion Public Methods

        #region Private Methods

        private EditableCollection<SlideEx> GetSlides()
        {
            var sdkPrePart = _outerSdkPresentation.PresentationPart;
            _preData = new PresentationData(sdkPrePart.Presentation, _slideSize);
            var slideCollection = SlideCollection.Create(sdkPrePart, _preData, this);

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
            _slides = new Lazy<EditableCollection<SlideEx>>(GetSlides);
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