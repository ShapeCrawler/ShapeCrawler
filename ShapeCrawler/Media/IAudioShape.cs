using DocumentFormat.OpenXml;
using ShapeCrawler.Media;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents an audio shape.
    /// </summary>
    public interface IAudioShape : IShape
    {
        /// <summary>
        ///     Gets bytes of audio content.
        /// </summary>
        public byte[] BinaryData { get; }

        /// <summary>
        ///     Gets MIME type.
        /// </summary>
        string MIME { get; }
    }

    internal class AudioShape : MediaShape, IAudioShape
    {
        internal AudioShape(OpenXmlCompositeElement pShapeTreesChild, SCSlide slide)
            : base(pShapeTreesChild, slide, null)
        {
        }

        public SCShapeType ShapeType => SCShapeType.AudioShape;
    }
}