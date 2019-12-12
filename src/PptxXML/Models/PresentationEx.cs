using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ObjectEx.Extensions;
using ObjectEx.Utilities;
using PptxXML.Models.Elements;
using PptxXML.Services;

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

            _pptxFileStream = new MemoryStream(pptxFileBytes);
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
            _slides = new SlideCollection(_xmlDoc);
            var sldNumber = 0;
            var groupShapeTypeParser = new GroupShapeTypeParser();
            var elementCreator = new ElementFactory(groupShapeTypeParser);
            var groupBuilder = new GroupEx.Builder(groupShapeTypeParser, elementCreator);
            foreach (var sldPart in _xmlDoc.PresentationPart.SlideParts)
            {
                sldNumber++;
                _slides.Add(new SlideEx(sldPart, sldNumber, elementCreator, groupShapeTypeParser, groupBuilder));
            }
        }

        #endregion Private Methods
    }
}
