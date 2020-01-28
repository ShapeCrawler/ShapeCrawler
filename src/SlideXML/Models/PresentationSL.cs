using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using LogicNull.Extensions;
using LogicNull.Utilities;
using SlideXML.Extensions;
using SlideXML.Models.Settings;
using SlideXML.Services;

namespace SlideXML.Models
{
    /// <summary>
    /// Represents a presentation.
    /// </summary>
    public class PresentationSL : IPresentation
    {
        #region Fields

        private PresentationDocument _xmlDoc;
        private ISlideCollection _slides;
        private bool _disposed;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets slides.
        /// </summary>
        public ISlideCollection Slides
        {
            get
            {
                if (_slides == null)
                {
                    InitSlides();
                }

                return _slides;
            }
        }

        /// <summary>
        /// Returns presentation slides width in EMUs.
        /// </summary>
        public int SlideWidth => _xmlDoc.PresentationPart.Presentation.SlideSize.Cx.Value;

        /// <summary>
        /// Returns presentation slides height in EMUs.
        /// </summary>
        public int SlideHeight => _xmlDoc.PresentationPart.Presentation.SlideSize.Cy.Value;

        #endregion Properties

        #region Constructors and Finalizer

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationSL"/> class by pptx-file stream.
        /// </summary>
        /// <param name="pptxFileStream"></param>
        public PresentationSL(Stream pptxFileStream)
        {
            pptxFileStream.ThrowIfNull(nameof(pptxFileStream));
            if (pptxFileStream.CanSeek)
            {
                pptxFileStream.Seek(0, SeekOrigin.Begin);
            }
            _xmlDoc = PresentationDocument.Open(pptxFileStream, true);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationSL"/> class by pptx-file byte array.
        /// </summary>
        /// <param name="pptxFileBytes"></param>
        public PresentationSL(byte[] pptxFileBytes)
        {
            Check.NotNull(pptxFileBytes, nameof(pptxFileBytes));
            var pptxMemoryStream = new MemoryStream();
            pptxMemoryStream.Write(pptxFileBytes, 0, pptxFileBytes.Length);
            _xmlDoc = PresentationDocument.Open(pptxMemoryStream, true);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationSL"/> class by pptx-file path.
        /// </summary>
        public PresentationSL(string pptxFilePath)
        {
            Check.NotEmpty(pptxFilePath, nameof(pptxFilePath));            
            _xmlDoc = PresentationDocument.Open(pptxFilePath, true);
        }

        /// <summary>
        /// The Finalizer.
        /// </summary>
        ~PresentationSL()
        {
            DisposeManaged();
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Saves the presentation in specified file path. After saved the presentation is not closed.
        /// </summary>
        /// <param name="filePath"></param>
        public void SaveAs(string filePath)
        {
            Check.NotEmpty(filePath, nameof(filePath));
            _xmlDoc = (PresentationDocument)_xmlDoc.SaveAs(filePath);
        }

        /// <summary>
        /// Saves and closes the presentation.
        /// </summary>
        public void Close()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(GetType().Name);
            }
            DisposeManaged();
        }

        /// <summary>
        /// Saves and closes the presentation, and releases all resources.
        /// </summary>
        public void Dispose()
        {
            Close();
            GC.SuppressFinalize(this);
        }

        #endregion Public Methods

        #region Private Methods

        private void DisposeManaged()
        {
            _xmlDoc.Close();
            _disposed = true;
        }

        private void InitSlides()
        {
            PresentationPart presentationPart = _xmlDoc.PresentationPart;
            var nbSlides = presentationPart.SlideParts.Count();
            _slides = new SlideCollection(_xmlDoc, nbSlides);
            var groupShapeTypeParser = new GroupShapeTypeParser(); // TODO: inject via DI Container
            var bgImgFactory = new BackgroundImageFactory();
            var preSettings = new PreSettings(_xmlDoc.PresentationPart.Presentation);

            for (var slideIndex = 0; slideIndex < nbSlides; slideIndex++)
            {
                SlidePart slidePart = presentationPart.GetSlidePartByIndex(slideIndex);
                var newSldEx = new SlideSL(slidePart, 
                                   slideIndex + 1,
                                   groupShapeTypeParser,
                                   bgImgFactory,
                                   preSettings);
                _slides.Add(newSldEx);
            }
        }

        #endregion Private Methods
    }
}