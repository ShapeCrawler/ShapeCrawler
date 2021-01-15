using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Factories.ShapeCreators;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Settings;

namespace ShapeCrawler.Collections
{
    /// <summary>
    /// Represents a collection of the slide shapes.
    /// </summary>
    public class ShapesCollection : LibraryCollection<ShapeSc>
    {
        #region Constructors

        public ShapesCollection(SlidePart sdkSldPart, IPresentationData preSettings, SlideSc slideEx) :
            this(sdkSldPart, new ShapeFactory(preSettings), slideEx)
        {
            
        }

        public ShapesCollection(SlidePart sdkSldPart, ShapeFactory shapeFactory, SlideSc slideEx)
        {
            CollectionItems = shapeFactory.FromSdlSlidePart(sdkSldPart, slideEx);
        }

        #endregion Constructors
    }
}