using System.IO;
using DocumentFormat.OpenXml.Packaging;
using LogicNull.Utilities;
using ObjectEx.Extensions;

namespace SlideXML.Models
{
    /// <summary>
    /// Represents a image model.
    /// </summary>
    public class ImageEx
    {
        #region Fields

        private readonly SlidePart _sldPart;
        private ImagePart _imgPart;
        private byte[] _bytes;
        private readonly string _blipRelateId;

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

        public ImageEx(SlidePart sldPart, string blipRelateId)
        {
            Check.NotNull(sldPart, nameof(sldPart));
            Check.NotNull(blipRelateId, nameof(blipRelateId));
            _sldPart = sldPart;
            _blipRelateId = blipRelateId;
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

            _bytes = null; // reset
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
            if (_imgPart == null) //TODO: use ?? operator
            {
                _imgPart = (ImagePart)_sldPart.GetPartById(_blipRelateId);
            }

            return _imgPart;
        }

        #endregion Private Methods
    }
}