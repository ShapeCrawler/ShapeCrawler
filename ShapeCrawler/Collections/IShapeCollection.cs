using System.Collections.Generic;
using System.IO;
using ShapeCrawler.Audio;
using ShapeCrawler.Shapes;

namespace ShapeCrawler.Collections
{
    public interface IShapeCollection : IEnumerable<IShape>
    {
        /// <summary>
        ///     Gets the element at the specified index.
        /// </summary>
        IShape this[int index] { get; }

        /// <summary>
        ///     Gets the number of series items in the collection.
        /// </summary>
        int Count { get; }

        /// <summary>
        ///     Gets the generic enumerator that iterates through the collection.
        /// </summary>
        IEnumerator<IShape> GetEnumerator();

        IAudioShape AddNewAudio(int x, int y, Stream audioStream);
    }
}