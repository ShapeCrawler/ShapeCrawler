using ShapeCrawler.Exceptions;
using ShapeCrawler.Placeholders;

namespace ShapeCrawler.Factories
{
    /// <summary>
    ///     <inheritdoc cref="ILocation" />
    /// </summary>
    internal class PlaceholderTransform : ILocation
    {
        private readonly PlaceholderLocationData _placeholderLocationData;

        #region Constructors

        public PlaceholderTransform(PlaceholderLocationData placeholderLocationData)
        {
            _placeholderLocationData = placeholderLocationData;
        }

        #endregion Constructors

        public int X => _placeholderLocationData.X;

        public int Y => _placeholderLocationData.Y;

        public int Width => _placeholderLocationData.Width;

        public int Height => _placeholderLocationData.Height;

        #region Public Methods

        public void SetX(int x)
        {
            throw new NotSupportedFeature(ExceptionMessages.PropertyCanChangedInNextVersion);
        }

        public void SetY(int y)
        {
            throw new NotSupportedFeature(ExceptionMessages.PropertyCanChangedInNextVersion);
        }

        public void SetWidth(int w)
        {
            throw new NotSupportedFeature(ExceptionMessages.PropertyCanChangedInNextVersion);
        }

        public void SetHeight(int h)
        {
            throw new NotSupportedFeature(ExceptionMessages.PropertyCanChangedInNextVersion);
        }

        #endregion Public Methods
    }
}