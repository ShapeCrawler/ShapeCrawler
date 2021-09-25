using DocumentFormat.OpenXml;

namespace ShapeCrawler.Audio
{
    internal class AudioShape : SlideShape, IAudioShape
    {
        public AudioShape(OpenXmlCompositeElement pShapeTreesChild, SCSlide parentSlide)
            : base(pShapeTreesChild, parentSlide, null)
        {
        }

        public byte[] BinaryData { get; }
    }
}