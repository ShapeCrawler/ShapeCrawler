using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using ObjectEx.Extensions;
using ObjectEx.Utilities;
using PptxXML.Enums;
using PptxXML.Exceptions;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxXML.Models.Elements
{
    /// <summary>
    /// Represents a picture element.
    /// </summary>
    public class PictureEx: Element
    {
        #region Fields

        private readonly SlidePart _sldPart;
        private ImagePart _imgPart;
        private byte[] _bytes;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Gets image bytes.
        /// </summary>
        /// <returns>
        /// A <c>byte array</c>, otherwise <c>null</c> if image is not exist.
        /// </returns>
        public byte[] Bytes
        {
            get
            {
                if (_bytes == null)
                {
                    InitBytes();
                }

                return _bytes;
            }
        }

        #endregion Properties

        #region Constructors

        /// <summary>
        /// Initializes a new instance of <see cref="PictureEx"/> class.
        /// </summary>
        public PictureEx(SlidePart sldPart) : base(ElementType.Picture)
        {
            Check.NotNull(sldPart, nameof(sldPart));
            _sldPart = sldPart;
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

            var imgPart = GetImagePart();
            stream.SeekBegin();
            imgPart.FeedData(stream);
            
            _bytes = null;
        }

        #endregion Public Methods

        #region Private Methods

        private void InitBytes()
        {
            var imgPart = GetImagePart();

            using (var stream = imgPart.GetStream())
            {
                var length = stream.Length;
                _bytes = new byte[length];
                stream.Read(_bytes, 0, (int)stream.Length); //TODO: use stream.ReadAsync instead
            }
        }

        private ImagePart GetImagePart()
        {
            if (_imgPart == null)
            {
                // imagePart
                var pic = (P.Picture) XmlCompositeElement;
                var pBlipFill = pic.GetFirstChild<P.BlipFill>();
                var picEmbedValue = pBlipFill?.Blip?.Embed?.Value;
                if (picEmbedValue == null)
                {
                    throw new PptxXMLException("Element does contain an image.");
                }

                _imgPart = (ImagePart) _sldPart.GetPartById(picEmbedValue);
            }

            return _imgPart;
        }

        #endregion Private Methods
    }
}