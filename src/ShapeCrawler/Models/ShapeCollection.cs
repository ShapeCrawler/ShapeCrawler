using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Models.Settings;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Services.ShapeCreators;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;
// ReSharper disable All

namespace SlideDotNet.Models
{
    /// <summary>
    /// Represents a collection of the slide shapes.
    /// </summary>
    public class ShapeCollection : LibraryCollection<ShapeEx>
    {
        #region Constructors

        /// <summary>
        /// Initializes a new instance by default <see cref="ShapeFactory"/> instance.
        /// </summary>
        /// <param name="sdkSldPart"></param>
        /// <param name="preSettings"></param>
        public ShapeCollection(SlidePart sdkSldPart, IPreSettings preSettings):
            this(sdkSldPart, new ShapeFactory(preSettings))
        {
            
        }

        public ShapeCollection(SlidePart sdkSldPart, IShapeFactory shapeFactory)
        {
            Check.NotNull(sdkSldPart, nameof(sdkSldPart));
            Check.NotNull(shapeFactory, nameof(shapeFactory));

            CollectionItems = shapeFactory.FromSldPart(sdkSldPart).ToList();
        }

        #endregion Constructors
    }
}