using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a media content.
/// </summary>
public interface IMedia
{
    /// <summary>
    ///     Gets MIME type.
    /// </summary>
    // ReSharper disable once InconsistentNaming
    string MIME { get; }

    /// <summary>
    ///     Gets bytes of video content.
    /// </summary>
    public byte[] AsByteArray();

#if DEBUG
    /// <summary>
    ///     Gets or sets audio start mode.
    /// </summary>
    AudioStartMode StartMode { get; set; }
#endif
}

internal class Media(SlideShapeOutline outline, ShapeFill fill, P.Picture pPicture) : IMedia
{
    public IShapeOutline Outline => outline;

    public IShapeFill Fill => fill;

    public string MIME
    {
        get
        {
            var openXmlPart = pPicture.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
            var p14Media = pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!
                .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
            var relationship =
                openXmlPart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);

            return relationship.DataPart.ContentType;
        }
    }

    public AudioStartMode StartMode
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    public byte[] AsByteArray()
    {
        var openXmlPart = pPicture.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var p14Media = pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!
            .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single();
        var relationship = openXmlPart.DataPartReferenceRelationships.First(r => r.Id == p14Media.Embed!.Value);
        var stream = relationship.DataPart.GetStream();
        var ms = new MemoryStream();
        stream.CopyTo(ms);
        stream.Close();

        return ms.ToArray();
    }

    public void SetVideo(Stream video)
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
        var openXmlPart = pPicture.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;

        // The <p14:media> element stores a relationship ID pointing to the media data part
        var p14Media = pPicture.NonVisualPictureProperties!
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