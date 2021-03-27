using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shared;

namespace ShapeCrawler.Drawing
{
    /// <summary>
    ///     Represents an image model.
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class SCImage
    {
        #region Fields

        private readonly SlidePart _slidePart;
        private ImagePart _imagePart;
        private byte[] _bytes;
        private readonly string _blipRelateId;

        #endregion Fields

        #region Constructors

        public SCImage(SlidePart sdkSlidePart, string blipRelateId)
        {
            _slidePart = sdkSlidePart ?? throw new ArgumentNullException(nameof(sdkSlidePart));
            _blipRelateId = blipRelateId ?? throw new ArgumentNullException(nameof(blipRelateId));
        }

        #endregion Constructors

        #region Public Methods

#if NET5_0 || NETSTANDARD2_1 || NETCOREAPP2_1
        public async ValueTask<byte[]> GetImageBytes()
        {
            if (_bytes != null)
            {
                return _bytes; // return from cache
            }

            using var imgPartStream = GetImagePart().GetStream();
            _bytes = new byte[imgPartStream.Length];
            await imgPartStream.ReadAsync(_bytes.AsMemory(0, (int) imgPartStream.Length)).ConfigureAwait(false);

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
            await imgPartStream.ReadAsync(_bytes, 0, (int) imgPartStream.Length).ConfigureAwait(false);
            ;
            imgPartStream.Close();
            return _bytes;
        }
#endif

        /// <summary>
        ///     Sets an image.
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
            return _imagePart ??= (ImagePart) _slidePart.GetPartById(_blipRelateId);
        }

        #endregion Private Methods
    }
}