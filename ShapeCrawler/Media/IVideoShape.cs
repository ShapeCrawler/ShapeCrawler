using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;

namespace ShapeCrawler.Media
{
    /// <summary>
    ///     Represents a shape containing video content.
    /// </summary>
    public interface IVideoShape : IShape
    {
        /// <summary>
        ///     Gets bytes of video content.
        /// </summary>
        public byte[] BinaryData { get; }
        
        /// <summary>
        ///     Gets MIME type.
        /// </summary>
        string MIME { get; }
    }

    internal class VideoShape : MediaShape, IVideoShape
    {
        internal VideoShape(SCSlide slide, OpenXmlCompositeElement pShapeTreeChild)
            : base(pShapeTreeChild, slide, null)
        {
        }

        public SCShapeType ShapeType => SCShapeType.VideoShape;
    }
}