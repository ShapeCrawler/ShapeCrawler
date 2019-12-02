using PptxXML.Entities;
using System.Collections.Generic;

namespace PptxXML.Models
{
    /// <summary>
    /// Provides APIs for slide collection.
    /// </summary>
    public interface ISlideCollection: IEnumerable<SlideEx>
    {
        void Add(SlideEx item);

        void Remove(SlideEx item);
    }
}