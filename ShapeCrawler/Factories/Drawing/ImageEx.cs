using System;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shared;

namespace ShapeCrawler.Factories.Drawing
{
    /// <summary>
    /// Represents an image model.
    /// </summary>
    public class ImageEx
    {
        #region Fields

        private readonly SlidePart _sldPart;
        private ImagePart _imgPart;
        private byte[] _bytes;
        private readonly string _blipRelateId;

        #endregion Fields

        #region Constructors

        public ImageEx(SlidePart sdkSlidePart, string blipRelateId)
        {
            _sldPart = sdkSlidePart ?? throw new ArgumentNullException(nameof(sdkSlidePart));
            _blipRelateId = blipRelateId ?? throw new ArgumentNullException(nameof(blipRelateId));
        }

        #endregion Constructors

        #region Public Methods

#if NETSTANDARD2_1 || NETCOREAPP2_0
        public async ValueTask<byte[]> GetImageBytes()
        {
            if (_bytes != null)
            {
                return _bytes; // return from cache
            }

            using var imgPartStream = GetImagePart().GetStream();
            _bytes = new byte[imgPartStream.Length];
            await imgPartStream.ReadAsync(_bytes, 0, (int)imgPartStream.Length).ConfigureAwait(false);

            return _bytes;
        }
#else
        public async Task<byte[]> GetImageBytes()
        {
            if (_bytes != null)
            {
                return _bytes; // return from cache
            }
            var imgPartStream = GetImagePart().GetStream();
            _bytes = new byte[imgPartStream.Length];
            await imgPartStream.ReadAsync(_bytes, 0, (int)imgPartStream.Length).ConfigureAwait(false);;

            return _bytes;
        }
#endif

        /// <summary>
        /// Sets an image.
        /// </summary>
        /// <param name="stream"></param>
        public void SetImage(Stream stream)
        {
            Check.NotNull(stream, nameof(stream));
            GetImagePart().FeedData(stream);
            _bytes = null; // resets cache
        }

#endregion Public Methods

        #region Private Methods

        private ImagePart GetImagePart()
        {
            return _imgPart ??= (ImagePart) _sldPart.GetPartById(_blipRelateId);
        }

        #endregion Private Methods
    }
}