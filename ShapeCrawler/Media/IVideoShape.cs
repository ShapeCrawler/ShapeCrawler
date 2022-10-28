using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using OneOf;

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
        internal VideoShape(OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide, OpenXmlCompositeElement pShapeTreeChild)
            : base(pShapeTreeChild, oneOfSlide, null)
        {
        }

        public override SCShapeType ShapeType => SCShapeType.VideoShape;
    }
}