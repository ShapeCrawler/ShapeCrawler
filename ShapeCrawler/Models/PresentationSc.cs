using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Models;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using ShapeCrawler.Statics;
// ReSharper disable CheckNamespace

namespace ShapeCrawler
{
    [SuppressMessage("ReSharper", "SuggestVarOrType_Elsewhere")]
    public class PresentationSc
    {
        #region Fields
        // TODO: Implement IDisposable
        private PresentationDocument _presentationDocument;
        private Lazy<EditableCollection<SlideSc>> _slides;
        private Lazy<SlideSizeSc> _slideSize;
        private bool _closed;
        private PresentationData _preData;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets the presentation slides.
        /// </summary>
        public EditableCollection<SlideSc> Slides => _slides.Value;

        /// <summary>
        /// Gets the presentation slides width.
        /// </summary>
        public int SlideWidth => _slideSize.Value.Width;

        /// <summary>
        /// Gets the presentation slides height.
        /// </summary>
        public int SlideHeight => _slideSize.Value.Height;

        public SlideMasterCollection SlideMasters => SlideMasterCollection.Create(_presentationDocument.PresentationPart.SlideMasterParts);

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationSc"/> class by pptx-file path.
        /// </summary>
        internal PresentationSc(string pptxPath, bool isEditable)
        {
            ThrowIfSourceInvalid(pptxPath);
            _presentationDocument = PresentationDocument.Open(pptxPath, isEditable);
            Init();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationSc"/> class by pptx-file stream.
        /// </summary>
        public PresentationSc(Stream pptxStream, bool isEditable = false)
        {
            ThrowIfSourceInvalid(pptxStream);
            _presentationDocument = PresentationDocument.Open(pptxStream, isEditable);
            Init();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationSc"/> class by pptx-file stream.
        /// </summary>
        private PresentationSc(MemoryStream pptxStream, bool isEditable)
        {
            ThrowIfSourceInvalid(pptxStream);
            _presentationDocument = PresentationDocument.Open(pptxStream, isEditable);
            Init();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationSc"/> class by pptx-file byte array.
        /// </summary>
        /// <param name="pptxBytes"></param>
        public PresentationSc(byte[] pptxBytes)
        {
            ThrowIfSourceInvalid(pptxBytes);

            var pptxStream = new MemoryStream();
            pptxStream.Write(pptxBytes, 0, pptxBytes.Length);
            _presentationDocument = PresentationDocument.Open(pptxStream, true);
            
            Init();
        }

        #endregion Constructors

        #region Public Methods

        public static PresentationSc Open(string pptxPath, bool isEditable)
        {
            return new PresentationSc(pptxPath, isEditable);
        }

        public void Save()
        {
            _presentationDocument.Save();
        }

        /// <summary>
        /// Saves presentation in specified file path.
        /// </summary>
        /// <param name="filePath"></param>
        public void SaveAs(string filePath)
        {
            Check.NotEmpty(filePath, nameof(filePath));
            _presentationDocument = (PresentationDocument)_presentationDocument.SaveAs(filePath);
        }

        /// <summary>
        /// Saves presentation in specified stream.
        /// </summary>
        /// <param name="stream"></param>
        public void SaveAs(Stream stream)
        {
            Check.NotNull(stream, nameof(stream));
            _presentationDocument = (PresentationDocument)_presentationDocument.Clone(stream);
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

            _presentationDocument.Close();
            if (_preData != null)
            {
                foreach (var xlsxDoc in _preData.XlsxDocuments.Values)
                {
                    xlsxDoc.Close();
                }
            }

            _closed = true;
        }

        public static PresentationSc Open(byte[] pptxBytes, bool isEditable = false)
        {
            ThrowIfSourceInvalid(pptxBytes);

            var pptxMemoryStream = new MemoryStream();
            pptxMemoryStream.Write(pptxBytes, 0, pptxBytes.Length);

            return new PresentationSc(pptxMemoryStream, isEditable);
        }

        public static PresentationSc Open(Stream stream, bool isEditable)
        {
            return new PresentationSc(stream, isEditable);
        }

        #endregion Public Methods

        #region Private Methods

        private EditableCollection<SlideSc> GetSlides()
        {
            var sdkPrePart = _presentationDocument.PresentationPart;
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
            _slides = new Lazy<EditableCollection<SlideSc>>(GetSlides);
            _slideSize = new Lazy<SlideSizeSc>(ParseSlideSize);
        }

        private SlideSizeSc ParseSlideSize()
        {
            var sdkSldSize = _presentationDocument.PresentationPart.Presentation.SlideSize;
            return new SlideSizeSc(sdkSldSize.Cx.Value, sdkSldSize.Cy.Value);
        }

        #endregion Private Methods
    }
}