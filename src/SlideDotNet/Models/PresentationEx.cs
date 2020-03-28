using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Collections;
using SlideDotNet.Exceptions;
using SlideDotNet.Models.Settings;
using SlideDotNet.Statics;
using SlideDotNet.Validation;

namespace SlideDotNet.Models
{
    /// <summary>
    /// <inheritdoc cref="IPresentation"/>
    /// </summary>
    public class PresentationEx : IPresentation
    {
        #region Fields

        private PresentationDocument _sdkPre;
        private readonly Lazy<EditAbleCollection<Slide>> _slides;
        private bool _disposed;

        #endregion Fields

        #region Properties

        public EditAbleCollection<Slide> Slides => _slides.Value;

        public int SlideWidth => _sdkPre.PresentationPart.Presentation.SlideSize.Cx.Value;

        public int SlideHeight => _sdkPre.PresentationPart.Presentation.SlideSize.Cy.Value;

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationEx"/> class by pptx-file path.
        /// </summary>
        public PresentationEx(string pptxPath)
        {
            ThrowIfInvalid(pptxPath);
            _sdkPre = PresentationDocument.Open(pptxPath, true);
            ThrowIfSlidesNumberLarge();
            _slides = new Lazy<EditAbleCollection<Slide>>(InitSlides);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationEx"/> class by pptx-file stream.
        /// </summary>
        /// <param name="pptxStream"></param>
        public PresentationEx(Stream pptxStream)
        {
            ThrowIfInvalid(pptxStream);
            _sdkPre = PresentationDocument.Open(pptxStream, true);
            ThrowIfSlidesNumberLarge();
            _slides = new Lazy<EditAbleCollection<Slide>>(InitSlides);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationEx"/> class by pptx-file byte array.
        /// </summary>
        /// <param name="pptxBytes"></param>
        public PresentationEx(byte[] pptxBytes)
        {
            ThrowIfInvalid(pptxBytes);
            var pptxStream = new MemoryStream();
            pptxStream.Write(pptxBytes, 0, pptxBytes.Length);
            _sdkPre = PresentationDocument.Open(pptxStream, true);
            ThrowIfSlidesNumberLarge();
            _slides = new Lazy<EditAbleCollection<Slide>>(InitSlides);
        }

        #endregion Constructors

        #region Public Methods

        public void SaveAs(string filePath)
        {
            Check.NotEmpty(filePath, nameof(filePath));
            _sdkPre = (PresentationDocument)_sdkPre.SaveAs(filePath);
        }

        public void SaveAs(Stream stream)
        {
            Check.NotNull(stream, nameof(stream));
            _sdkPre = (PresentationDocument)_sdkPre.Clone(stream);
        }

        public void Close()
        {
            if (_disposed)
            {
                return;
            }
            _sdkPre.Close();
            _disposed = true;

            
        }

        public void Dispose()
        {
            Close();
        }

        #endregion Public Methods

        #region Private Methods

        private EditAbleCollection<Slide> InitSlides()
        {
            var preSettings = new PreSettings(_sdkPre.PresentationPart.Presentation);
            var slideCollection = SlideCollection.Create(_sdkPre, preSettings);

            return slideCollection;
        }

        private void ThrowIfInvalid(string path)
        {
            if (!File.Exists(path))
            {
                throw new FileNotFoundException(nameof(path));
            }
            var  fileInfo = new FileInfo(path);

            ThrowIfPptxSizeLarge(fileInfo.Length);
        }

        private void ThrowIfInvalid(Stream stream)
        {
            Check.NotNull(stream, nameof(stream));

            ThrowIfPptxSizeLarge(stream.Length);
        }

        private void ThrowIfInvalid(byte[] bytes)
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

        #endregion Private Methods
    }
}