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

        private PresentationDocument _xmlDoc;
        private readonly Lazy<ISlideCollection> _slides;
        private bool _disposed;

        #endregion Fields

        #region Properties

        /// <summary>
        /// <inheritdoc cref="IPresentation.Slides"/>
        /// </summary>
        public ISlideCollection Slides => _slides.Value;

        /// <summary>
        /// <inheritdoc cref="IPresentation.SlideWidth"/>
        /// </summary>
        public int SlideWidth => _xmlDoc.PresentationPart.Presentation.SlideSize.Cx.Value;

        /// <summary>
        /// <inheritdoc cref="IPresentation.SlideHeight"/>
        /// </summary>
        public int SlideHeight => _xmlDoc.PresentationPart.Presentation.SlideSize.Cy.Value;

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationEx"/> class by pptx-file path.
        /// </summary>
        public PresentationEx(string pptxPath)
        {
            ThrowIfInvalid(pptxPath);
            _xmlDoc = PresentationDocument.Open(pptxPath, true);
            ThrowIfSlidesNumberLarge();
            _slides = new Lazy<ISlideCollection>(InitSlides);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationEx"/> class by pptx-file stream.
        /// </summary>
        /// <param name="pptxStream"></param>
        public PresentationEx(Stream pptxStream)
        {
            ThrowIfInvalid(pptxStream);
            _xmlDoc = PresentationDocument.Open(pptxStream, true);
            ThrowIfSlidesNumberLarge();
            _slides = new Lazy<ISlideCollection>(InitSlides);
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
            _xmlDoc = PresentationDocument.Open(pptxStream, true);
            ThrowIfSlidesNumberLarge();
            _slides = new Lazy<ISlideCollection>(InitSlides);
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// <inheritdoc cref="IPresentation.SaveAs"/>
        /// </summary>
        /// <param name="filePath"></param>
        public void SaveAs(string filePath)
        {
            Check.NotEmpty(filePath, nameof(filePath));
            _xmlDoc = (PresentationDocument)_xmlDoc.SaveAs(filePath);
        }

        /// <summary>
        /// <inheritdoc cref="IPresentation.Close"/>
        /// </summary>
        public void Close()
        {
            if (_disposed)
            {
                return;
            }
            _xmlDoc.Close();
            _disposed = true;
        }

        /// <summary>
        /// Saves and closes the current presentation.
        /// </summary>
        public void Dispose()
        {
            Close();
        }

        #endregion Public Methods

        #region Private Methods

        private ISlideCollection InitSlides()
        {
            var preSettings = new PreSettings(_xmlDoc.PresentationPart.Presentation);
            var slideCollection = SlideCollection.Create(_xmlDoc, preSettings);

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
            var nbSlides = _xmlDoc.PresentationPart.SlideParts.Count();
            if (nbSlides > Limitations.MaxSlidesNumber)
            {
                Close();
                throw SlidesMuchMoreException.FromMax(Limitations.MaxSlidesNumber);
            }
        }

        #endregion Private Methods
    }
}