using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents an image model.
    /// </summary>
    public class SCImage // TODO: make internal?
    {
        private readonly SCPresentation parentPresentation;
        private readonly IRemovable imageContainer;
        private readonly StringValue picReference;
        private readonly OpenXmlPart openXmlPart;
        private byte[]? bytes;

        private SCImage(
            ImagePart imagePart,
            IRemovable imageContainer,
            StringValue picReference,
            OpenXmlPart openXmlPart)
        {
            this.ImagePart = imagePart;
            this.imageContainer = imageContainer;
            this.picReference = picReference;
            this.openXmlPart = openXmlPart;
            
            this.parentPresentation = ((IPresentationComponent)imageContainer).PresentationInternal;
            this.MIME = this.ImagePart.ContentType;
        }

        public string MIME { get; }

        internal ImagePart ImagePart { get; private set; }

#if NET5_0

        /// <summary>
        ///     Gets bytes content.
        /// </summary>
        public async ValueTask<byte[]> GetBytes()
        {
            var stream = this.ImagePart.GetStream();
            this.bytes = new byte[stream.Length];
            await stream.ReadAsync(this.bytes.AsMemory(0, (int)stream.Length)).ConfigureAwait(false);

            stream.Close();
            return this.bytes;
        }

#else
        public async Task<byte[]> GetBytes()
        {
            if (bytes != null)
            {
                return bytes; // return from cache
            }

            Stream stream = this.ImagePart.GetStream();
            bytes = new byte[stream.Length];
            await stream.ReadAsync(bytes, 0, (int) stream.Length).ConfigureAwait(false);
            stream.Close();
            return bytes;
        }
#endif

        /// <summary>
        ///     Sets image with stream.
        /// </summary>
        public void SetImage(Stream sourceStream)
        {
            this.imageContainer.ThrowIfRemoved();

            var isSharedImagePart = this.parentPresentation.ImageParts.Count(ip => ip == this.ImagePart) > 1;
            if (isSharedImagePart)
            {
                var rId = RelatedIdGenerator.Generate();
                this.ImagePart = this.openXmlPart.AddNewPart<ImagePart>("image/png", rId);
                this.picReference.Value = rId;
            }

            sourceStream.Position = 0;
            this.ImagePart.FeedData(sourceStream);
            this.bytes = null; // resets cache
        }

        /// <summary>
        ///     Sets image with byte array.
        /// </summary>
        public void SetImage(byte[] sourceBytes)
        {
            var stream = new MemoryStream();
            stream.Write(sourceBytes, 0, sourceBytes.Length);
            this.SetImage(stream);
        }

#if NETSTANDARD2_0
        public void SetImage(string filePath)
        {
            byte[] sourceBytes = File.ReadAllBytes(filePath);
            this.SetImage(sourceBytes);
        }
#else

        /// <summary>
        ///     Sets image from file contents.
        /// </summary>
        public async Task SetImage(string filePath)
        {
            byte[] sourceBytes = await File.ReadAllBytesAsync(filePath).ConfigureAwait(false);
            this.SetImage(sourceBytes);
        }
#endif

        internal static SCImage ForPicture(Shape pictureShape, OpenXmlPart openXmlPart, StringValue picReference)
        {
            var imagePart = (ImagePart)openXmlPart.GetPartById(picReference.Value);

            return new SCImage(imagePart, pictureShape, picReference, openXmlPart);
        }

        internal static SCImage? ForBackground(SCSlide slide)
        {
            var pBackground = slide.SDKSlidePart.Slide.CommonSlideData.Background;
            if (pBackground == null)
            {
                return null;
            }

            var aBlipFill = pBackground.Descendants<A.BlipFill>().SingleOrDefault();
            var picReference = aBlipFill?.Blip?.Embed;
            if (picReference == null)
            {
                return null;
            }

            var imagePart = (ImagePart)slide.SDKSlidePart.GetPartById(picReference.Value);
            var backgroundImage = new SCImage(imagePart, slide, picReference, slide.SDKSlidePart);

            return backgroundImage;
        }

        internal static SCImage? ForAutoShapeFill(Shape autoShape, SlidePart slidePart)
        {
            var pShape = (P.Shape)autoShape.PShapeTreesChild;
            var aBlipFill = pShape.ShapeProperties!.GetFirstChild<A.BlipFill>();

            var picReference = aBlipFill?.Blip?.Embed;
            if (picReference == null)
            {
                return null;
            }

            var imagePart = (ImagePart)slidePart.GetPartById(picReference.Value!);

            return new SCImage(imagePart, autoShape, picReference, slidePart);
        }

        internal static SCImage Create(ImagePart imagePart, MasterPicture masterPic, StringValue stringValue, SlideMasterPart sldMasterPart)
        {
            return new SCImage(imagePart, masterPic, stringValue, sldMasterPart);
        }

        internal static SCImage Create(ImagePart imagePart, LayoutPicture layoutPic, StringValue stringValue, SlideLayoutPart slideLayoutPart)
        {
            return new SCImage(imagePart, layoutPic, stringValue, slideLayoutPart);
        }
    }
}