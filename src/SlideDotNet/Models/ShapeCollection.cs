using System.Collections.Generic;
using SlideDotNet.Collections;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Validation;

namespace SlideDotNet.Models
{
    /// <summary>
    /// Represents a collection of the slide shapes.
    /// </summary>
    public class ShapeCollection : LibraryCollection<ShapeEx>
    {
        #region Constructors

        public ShapeCollection(IEnumerable<ShapeEx> shapes)
        {
            Check.NotEmpty(shapes, nameof(shapes));
            CollectionItems = new List<ShapeEx>(shapes);
        }

        #endregion Constructors
    }
}