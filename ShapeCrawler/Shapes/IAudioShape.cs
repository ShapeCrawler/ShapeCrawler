using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents an audio shape.
    /// </summary>
    public interface IAudioShape : IShape
    {
        /// <summary>
        ///     Gets audio's data in bytes.
        /// </summary>
        byte[] BinaryData { get; } // TODO: add setter

        // TODO: add ContentType property containing MIME type of audio
    }

    internal class AudioShape : SlideShape, IAudioShape
    {
        internal AudioShape(OpenXmlCompositeElement pShapeTreesChild, SCSlide parentSlideLayoutInternal)
            : base(pShapeTreesChild, parentSlideLayoutInternal, null)
        {
        }

        public byte[] BinaryData { get; }

        public ShapeType ShapeType => ShapeType.AudioShape;
    }
}
