using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ObjectEx.Extensions;
using ObjectEx.Utilities;
using PptxXML.Extensions;
using PptxXML.Models.Elements;
using PptxXML.Models.Settings;
using PptxXML.Models.TextBody;
using PptxXML.Services;
using PptxXML.Services.Placeholder;
using PptxXML.Services.Placeholders;

namespace PptxXML.Models
{
    /// <summary>
    /// Represents a presentation.
    /// </summary>
    public class PresentationEx : IPresentationEx
    {
        #region Fields

        private readonly PresentationDocument _xmlDoc;
        private readonly MemoryStream _pptxFileStream;
        private ISlideCollection _slides;

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
        /// Initializes a new instance of the <see cref="PresentationEx"/> class by pptx-file stream.
        /// </summary>
        /// <param name="pptxFileStream"></param>
        public PresentationEx(Stream pptxFileStream)
        {
            pptxFileStream.ThrowIfNull(nameof(pptxFileStream));
            if (pptxFileStream.CanSeek)
            {
                pptxFileStream.Seek(0, SeekOrigin.Begin);
            }
            _xmlDoc = PresentationDocument.Open(pptxFileStream, true);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationEx"/> class by pptx-file byte array.
        /// </summary>
        /// <param name="pptxFileBytes"></param>
        public PresentationEx(byte[] pptxFileBytes)
        {
            Check.NotNull(pptxFileBytes, nameof(pptxFileBytes));
            _pptxFileStream = new MemoryStream();
            _pptxFileStream.Write(pptxFileBytes, 0, pptxFileBytes.Length);
            _xmlDoc = PresentationDocument.Open(_pptxFileStream, true);
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Saves the presentation in specified file path.
        /// </summary>
        /// <param name="filePath"></param>
        public void SaveAs(string filePath)
        {
            Check.NotEmpty(filePath, nameof(filePath));

            var savedXmlDoc = _xmlDoc.SaveAs(filePath);
            savedXmlDoc.Dispose();
        }

        /// <summary>
        /// Saves and closes the presentation, and releases all resources.
        /// </summary>
        public void Dispose()
        {
            _xmlDoc.Dispose(); // saves and closes
            _pptxFileStream?.Dispose();
        }

        #endregion Public Methods

        #region Private Methods

        private void InitSlides()
        {
            PresentationPart presentationPart = _xmlDoc.PresentationPart;
            var nbSlides = presentationPart.SlideParts.Count();
            _slides = new SlideCollection(_xmlDoc, nbSlides);
            var groupShapeTypeParser = new GroupShapeTypeParser(); // TODO: inject via DI Container
            var sldLayoutPartParser = new SlideLayoutPartParser();
            var bgImgFactory = new BackgroundImageFactory();
            var preSettings = new PreSettings(_xmlDoc.PresentationPart.Presentation);
            var txtBodyBuilder = new TextBodyEx.TextBodyExBuilder(new ParagraphEx.ParagraphExBuilder());
            var elFactory = new ElementFactory(new ShapeEx.Builder(bgImgFactory, txtBodyBuilder));
            var groupExBuilder = new GroupEx.Builder(groupShapeTypeParser, elFactory);

            for (var slideIndex = 0; slideIndex < nbSlides; slideIndex++)
            {
                SlidePart slidePart = presentationPart.GetSlidePartByIndex(slideIndex);
                var newSldEx = new SlideEx(slidePart, 
                                   slideIndex + 1, 
                                           elFactory,
                                           groupShapeTypeParser, 
                                           groupExBuilder, 
                                           sldLayoutPartParser, 
                                           bgImgFactory,
                                           preSettings);
                _slides.Add(newSldEx);
            }
        }

        #endregion Private Methods
    }
}