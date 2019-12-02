using System.IO;
using DocumentFormat.OpenXml;
using objectEx.Extensions;
using PptxXML.Enums;

namespace PptxXML.Entities.Elements
{
    /// <summary>
    /// Represent a picture element
    /// </summary>
    public class PictureEx: Element
    {
        #region Constructors

        /// <summary>
        /// Initialise an instance of <see cref="PictureEx"/> class.
        /// </summary>
        /// <param name="xmlCompositeElement"></param>
        public PictureEx(OpenXmlCompositeElement xmlCompositeElement) :
            base(xmlCompositeElement)
        {
            Type = ElementType.Picture;
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Sets an image.
        /// </summary>
        /// <param name="stream"></param>
        public void SetImage(Stream stream)
        {
            stream.ThrowIfNull(nameof(stream));
        }

        #endregion
    }
}
