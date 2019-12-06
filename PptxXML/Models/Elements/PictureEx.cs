using System;
using System.IO;
using DocumentFormat.OpenXml;
using ObjectEx.Utilities;

namespace PptxXML.Models.Elements
{
    /// <summary>
    /// Represents a picture element.
    /// </summary>
    public class PictureEx: Element
    {
        #region Constructors

        /// <summary>
        /// Initialise an instance of <see cref="PictureEx"/> class.
        /// </summary>
        /// <param name="xmlCompositeElement"></param>
        public PictureEx(OpenXmlCompositeElement xmlCompositeElement) : base(xmlCompositeElement)
        {

        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Sets an image.
        /// </summary>
        /// <param name="stream"></param>
        public void SetImage(Stream stream)
        {
            Check.NotNull(stream, nameof(stream));
            throw new NotImplementedException();
        }

        #endregion Public Methods

        #region Private Methods


        #endregion Private Methods
    }
}