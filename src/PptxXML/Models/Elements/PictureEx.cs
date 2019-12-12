using System;
using System.IO;
using DocumentFormat.OpenXml;
using ObjectEx.Utilities;
using PptxXML.Enums;

namespace PptxXML.Models.Elements
{
    /// <summary>
    /// Represents a picture element.
    /// </summary>
    public class PictureEx: Element
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance of <see cref="PictureEx"/> class.
        /// </summary>
        public PictureEx() : base(ElementType.Picture) { }

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