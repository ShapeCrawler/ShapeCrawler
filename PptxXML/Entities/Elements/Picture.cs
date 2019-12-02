using System.IO;
using DocumentFormat.OpenXml;
using objectEx.Extensions;
using PptxXML.Enums;

namespace PptxXML.Entities.Elements
{
    /// <summary>
    /// Represent a picture element
    /// </summary>
    public class Picture: Element
    {
        #region Constructors

        /// <summary>
        /// Initialise an instance of <see cref="Picture"/> class.
        /// </summary>
        /// <param name="xmlCompositeElement"></param>
        public Picture(OpenXmlCompositeElement xmlCompositeElement) :
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
