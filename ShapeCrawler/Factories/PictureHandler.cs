using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Audio;
using ShapeCrawler.Drawing;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using ShapeCrawler.Video;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    /// <summary>
    ///     Represents a picture handler for p:pic and picture p:graphicFrame element.
    /// </summary>
    internal class PictureHandler : OpenXmlElementHandler
    {
        private readonly ShapeContext.Builder shapeContextBuilder;

        internal PictureHandler(ShapeContext.Builder shapeContextBuilder)
        {
            this.shapeContextBuilder = shapeContextBuilder ?? throw new ArgumentNullException(nameof(shapeContextBuilder));
        }

        public override IShape Create(OpenXmlCompositeElement pShapeTreeChild, SCSlide slide)
        {
            P.Picture pPicture;
            if (pShapeTreeChild is P.Picture treePic)
            {
                OpenXmlElement element = treePic.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties.ChildElements.FirstOrDefault();
                if (element is A.AudioFromFile)
                {
                    A.AudioFromFile aAudioFile = treePic.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties
                        .GetFirstChild<A.AudioFromFile>();

                    if (aAudioFile != null)
                    {
                        return new AudioShape(slide, pShapeTreeChild);
                    }
                }
                else if (element is A.VideoFromFile)
                {
                    A.VideoFromFile aVideoFile = (A.VideoFromFile)element;

                    if (aVideoFile != null)
                    {
                        return new VideoShape(slide, pShapeTreeChild);
                    }
                }

                pPicture = treePic;
            }
            else
            {
                pPicture = pShapeTreeChild.Descendants<P.Picture>().FirstOrDefault();
            }

            if (pPicture == null)
            {
                return this.Successor?.Create(pShapeTreeChild, slide);
            }

            StringValue picReference = pPicture.GetFirstChild<P.BlipFill>()?.Blip?.Embed;
            if (picReference == null)
            {
                return null;
            }

            ShapeContext spContext = this.shapeContextBuilder.Build(pShapeTreeChild);
            SlidePicture picture = new(slide, spContext, pPicture, picReference);

            return picture;
        }
    }
}