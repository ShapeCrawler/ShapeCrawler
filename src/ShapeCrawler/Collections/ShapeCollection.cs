using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Services.ShapeCreators;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;

namespace ShapeCrawler.Collections
{
    /// <summary>
    /// Represents a collection of the slide shapes.
    /// </summary>
    public class ShapeCollection : LibraryCollection<Shape>
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance by default <see cref="ShapeFactory"/> instance.
        /// </summary>
        /// <param name="sdkSldPart"></param>
        /// <param name="preSettings"></param>
        public ShapeCollection(SlidePart sdkSldPart, IPresentationData preSettings, Slide slide) :
            this(sdkSldPart, new ShapeFactory(preSettings), slide)
        {
            
        }

        public ShapeCollection(SlidePart sdkSldPart, IShapeFactory shapeFactory, Slide slide)
        {
            CollectionItems = shapeFactory.FromSdlSlidePart(sdkSldPart, slide).ToList(); // TODO: Check whether it is possible avoid ToList()
        }

        #endregion Constructors
    }
}