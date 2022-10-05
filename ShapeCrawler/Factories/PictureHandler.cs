using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Drawing;
using ShapeCrawler.Media;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    /// <summary>
    ///     Represents handler for p:pic and p:graphicFrame elements.
    /// </summary>
    internal class PictureHandler : OpenXmlElementHandler
    {
        internal override Shape? Create(OpenXmlCompositeElement compositeElementOfPShapeTree, SCSlide slide, SlideGroupShape groupShape)
        {
            P.Picture? pPicture;
            if (compositeElementOfPShapeTree is P.Picture treePic)
            {
                OpenXmlElement element = treePic.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties.ChildElements.FirstOrDefault();

                switch (element)
                {
                    case AudioFromFile:
                    {
                        var aAudioFile = treePic.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties
                            .GetFirstChild<A.AudioFromFile>();
                        if (aAudioFile is not null)
                        {
                            return new AudioShape(compositeElementOfPShapeTree, slide);
                        }

                        break;
                    }

                    case VideoFromFile file:
                    {
                        A.VideoFromFile aVideoFile = file;

                        if (aVideoFile != null)
                        {
                            return new VideoShape(slide, compositeElementOfPShapeTree);
                        }

                        break;
                    }
                }

                pPicture = treePic;
            }
            else
            {
                pPicture = compositeElementOfPShapeTree.Descendants<P.Picture>().FirstOrDefault();
            }

            if (pPicture == null)
            {
                return this.Successor?.Create(compositeElementOfPShapeTree, slide, groupShape);
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