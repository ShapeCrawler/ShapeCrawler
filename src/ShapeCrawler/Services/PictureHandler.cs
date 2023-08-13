using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using OneOf;
using ShapeCrawler.Charts;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Services;

internal sealed class PictureHandler
{
    private readonly List<ImagePart> imageParts;
    private readonly TypedOpenXmlPart slideTypedOpenXmlPart;

    public PictureHandler(List<ImagePart> imageParts, TypedOpenXmlPart slideTypedOpenXmlPart)
    {
        this.imageParts = imageParts;
        this.slideTypedOpenXmlPart = slideTypedOpenXmlPart;
    }

    internal SCShape? FromTreeChild(
        OpenXmlCompositeElement pShapeTreeChild,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideOf,
        OneOf<SCSlideShapes, SCSlideGroupShape> shapeCollectionOf,
        TypedOpenXmlPart slideTypedOpenXmlPart,
        List<ChartWorkbook> chartWorkbooks)
    {
        P.Picture? pPicture;
        if (pShapeTreeChild is P.Picture treePic)
        {
            var element = treePic.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!.ChildElements.FirstOrDefault();

            switch (element)
            {
                case AudioFromFile:
                {
                    var aAudioFile = treePic.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties
                        .GetFirstChild<A.AudioFromFile>();
                    if (aAudioFile is not null)
                    {
                        return new SCSlideAudio(pShapeTreeChild, slideOf, shapeCollectionOf, slideTypedOpenXmlPart);
                    }

                    break;
                }

                case VideoFromFile:
                {
                    return new SCSlideMediaShape(pShapeTreeChild, slideOf, shapeCollectionOf, this.slideTypedOpenXmlPart);
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
            return this.Successor?.FromTreeChild(pShapeTreeChild, slideOf, shapeCollectionOf, slideTypedOpenXmlPart);
        }

        var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
        var blipEmbed = aBlip?.Embed;
        if (blipEmbed is null)
        {
            return null;
        }

        var picture = new SCSlidePicture(pPicture, slideOf, shapeCollectionOf, aBlip!, this.slideTypedOpenXmlPart, this.imageParts);

        return picture;
    }
}