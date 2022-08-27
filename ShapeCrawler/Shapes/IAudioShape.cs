using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
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
        ///     Gets audio's data in bytes.
        /// </summary>
        byte[] BinaryData { get; }
    }

    internal class AudioShape : SlideShape, IAudioShape
    {
        private readonly AudioFromFile aAudioFile;

        internal AudioShape(OpenXmlCompositeElement pShapeTreesChild, SCSlide slide)
            : base(pShapeTreesChild, slide, null)
        {
        }

        public byte[] BinaryData => GetBinaryData();

        public ShapeType ShapeType => ShapeType.AudioShape;
        
        private byte[] GetBinaryData()
        {
            var pPic = (P.Picture)this.PShapeTreesChild;
            var p14Media = pPic.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!.Descendants<P14.Media>().Single();
            var relationship = this.Slide.SDKSlidePart.DataPartReferenceRelationships.First(r => r.Id == p14Media!.Embed!.Value);
            var stream = relationship.DataPart.GetStream();
            var bytes = stream.ToArray();
            stream.Close();

            return bytes;
        }
    }
}
