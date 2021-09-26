using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Audio;
using ShapeCrawler.Shapes;
using ShapeCrawler.Video;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    /// <summary>
    ///     Represents handler for p:pic and p:graphicFrame elements.
    /// </summary>
    internal class PictureHandler : OpenXmlElementHandler
    {
        public override IShape? Create(OpenXmlCompositeElement pShapeTreesChild, SCSlide slide, SlideGroupShape groupShape)
        {
            P.Picture? pPicture;
            if (pShapeTreesChild is P.Picture treePic)
            {
                OpenXmlElement element = treePic.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties.ChildElements.FirstOrDefault();

                if (element is A.AudioFromFile)
                {
                    A.AudioFromFile? aAudioFile = treePic.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties
                    .GetFirstChild<A.AudioFromFile>();
                    if (aAudioFile is not null)
                    {
                        return new AudioShape(pShapeTreesChild, slide);
                    }
                }
                else if(element is A.VideoFromFile)
                {
                    A.VideoFromFile aVideoFile = (A.VideoFromFile)element;

                    if (aVideoFile != null)
                    {
                        return new VideoShape(slide, pShapeTreesChild);
                    }
                }

                pPicture = treePic;
            }
            else
            {
                pPicture = pShapeTreesChild.Descendants<P.Picture>().FirstOrDefault();
            }

            if (pPicture == null)
            {
                return this.Successor?.Create(pShapeTreesChild, slide, groupShape);
            }

            StringValue? picReference = pPicture.GetFirstChild<P.BlipFill>()?.Blip?.Embed;
            if (picReference is null)
            {
                return null;
            }

            SlidePicture picture = new (pPicture, slide, picReference);

            return picture;
        }
    }
}