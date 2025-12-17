using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;
using Position = ShapeCrawler.Positions.Position;

namespace ShapeCrawler.MediaContent;

internal sealed class MediaShape : DrawingShape
{
    private readonly P.Picture pPicture;

    internal MediaShape(Position position, ShapeSize shapeSize, ShapeId shapeId, P.Picture pPicture)
        : base(position, shapeSize, shapeId, pPicture)
    {
        this.pPicture = pPicture;
        this.Media = new Media(new SlideShapeOutline(pPicture.ShapeProperties!), new ShapeFill(pPicture.ShapeProperties!), pPicture);
    }

    public override IMedia? Media { get; }

    public override void SetVideo(Stream video)
    {
        if (video is null)
        {
            throw new ArgumentNullException(nameof(video));
        }

        // Reset incoming stream position to ensure full copy
        if (video.CanSeek)
        {
            video.Position = 0;
        }

        // Locate the Open XML part that contains this picture
        var openXmlPart = this.pPicture.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;

        // The <p14:media> element stores a relationship ID pointing to the media data part
        var p14Media = this.pPicture.NonVisualPictureProperties!
            .ApplicationNonVisualDrawingProperties!
            .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>()
            .Single();

        var embedId = p14Media.Embed!.Value!;

        // Find the relationship on the containing part that matches this ID
        var relationship = openXmlPart.DataPartReferenceRelationships.First(r => r.Id == embedId);

        // Feed the new video data into the existing media data part
        video.Position = 0;
        relationship.DataPart.FeedData(video);
    }
}