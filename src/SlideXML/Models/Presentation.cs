using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using SlideXML.Extensions;
using SlideXML.Models.Settings;
using SlideXML.Services;
using SlideXML.Validation;

namespace SlideXML.Models
{
    /// <summary>
    /// Represents a presentation.
    /// </summary>
    public class Presentation : IPresentation
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

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Presentation"/> class by pptx-file stream.
        /// </summary>
        /// <param name="pptxFileStream"></param>
        public Presentation(Stream pptxFileStream)
        {
            Check.NotNull(pptxFileStream, nameof(pptxFileStream));
            pptxFileStream.SeekBegin();
            _xmlDoc = PresentationDocument.Open(pptxFileStream, true);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Presentation"/> class by pptx-file byte array.
        /// </summary>
        /// <param name="pptxFileBytes"></param>
        public Presentation(byte[] pptxFileBytes)
        {
            Check.NotNull(pptxFileBytes, nameof(pptxFileBytes));
            var pptxMemoryStream = new MemoryStream();
            pptxMemoryStream.Write(pptxFileBytes, 0, pptxFileBytes.Length);
            _xmlDoc = PresentationDocument.Open(pptxMemoryStream, true);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Presentation"/> class by pptx-file path.
        /// </summary>
        public Presentation(string pptxFilePath)
        {
            Check.NotEmpty(pptxFilePath, nameof(pptxFilePath));            
            _xmlDoc = PresentationDocument.Open(pptxFilePath, true);
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
        /// Saves and closes the current presentation if it is not already closed.
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
                var newSldEx = new Slide(slidePart, 
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