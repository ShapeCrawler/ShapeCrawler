﻿using System;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Shared;

namespace SlideDotNet.Services.Drawing
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

        #region Constructors

        public ImageEx(SlidePart sldPart, string blipRelateId)
        {
            _sldPart = sldPart ?? throw new ArgumentNullException(nameof(sldPart));
            _blipRelateId = blipRelateId ?? throw new ArgumentNullException(nameof(blipRelateId));
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Returns image bytes.
        /// </summary>
        /// <returns></returns>
        public async Task<byte[]> GetImageBytes()
        {
            if (_bytes != null)
            {
                return _bytes; // return from cache
            }

            using var imgPartStream = GetImagePart().GetStream();
            _bytes = new byte[imgPartStream.Length];
            await imgPartStream.ReadAsync(_bytes, 0, (int)imgPartStream.Length);

            return _bytes;
        }

        /// <summary>
        /// Sets an image.
        /// </summary>
        /// <param name="stream"></param>
        public void SetImageStream(Stream stream)
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