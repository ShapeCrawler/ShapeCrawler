using SlideDotNet.Exceptions;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Services.Placeholders;
using A = DocumentFormat.OpenXml.Drawing;

namespace SlideDotNet.Models.Transforms
{
    /// <summary>
    /// <inheritdoc cref="ILocation"/>
    /// </summary>
    public class PlaceholderTransform : ILocation
    {
        private readonly PlaceholderLocationData _placeholderLocationData;

        public long X => _placeholderLocationData.X;

        public long Y => _placeholderLocationData.Y;

        public long Width => _placeholderLocationData.Width;

        public long Height => _placeholderLocationData.Height;

        #region Constructors

        public PlaceholderTransform(PlaceholderLocationData placeholderLocationData)
        {
            _placeholderLocationData = placeholderLocationData;
        }

        #endregion Constructors

        #region Public Methods

        public void SetX(long x)
        {
            throw new NextVersionFeatureException(ExceptionMessages.PropertyCanChangedInNextVersion);
        }

        public void SetY(long y)
        {
            throw new NextVersionFeatureException(ExceptionMessages.PropertyCanChangedInNextVersion);
        }

        public void SetWidth(long w)
        {
            throw new NextVersionFeatureException(ExceptionMessages.PropertyCanChangedInNextVersion);
        }

        public void SetHeight(long h)
        {
            throw new NextVersionFeatureException(ExceptionMessages.PropertyCanChangedInNextVersion);
        }

        #endregion Public Methods
    }
}