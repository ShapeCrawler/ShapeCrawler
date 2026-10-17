using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.MediaContent;
using ShapeCrawler.Slides;
using A = DocumentFormat.OpenXml.Drawing;
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
    ///     Gets bytes of media content.
    /// </summary>
    byte[] AsByteArray();

    /// <summary>
    ///     Replaces audio content with the specified MP3 or WAVE stream.
    /// </summary>
    /// <param name="media">Audio stream.</param>
    void Update(Stream media);

    /// <summary>
    ///     Gets or sets audio start mode.
    /// </summary>
    AudioStartMode StartMode { get; set; }
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
        get => new AudioTiming(pPicture).StartMode;
        set => new AudioTiming(pPicture).StartMode = value;
    }

    private OpenXmlPart OpenXmlPart => pPicture.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;

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

    public void Update(Stream media)
    {
        if (media is null)
        {
            throw new ArgumentNullException(nameof(media));
        }

        var audioFromFile = this.AudioFromFile();
        var p14Media = this.P14Media();
        var (audioReference, mediaReference) = this.ExistingRelationships(audioFromFile, p14Media);
        if (!ReferenceEquals(audioReference.DataPart, mediaReference.DataPart))
        {
            throw new SCException("Audio references must target the same media data part.");
        }

        using var bufferedMedia = this.BufferedMedia(media);
        var replacementDataPart = this.CreateMediaDataPart(bufferedMedia);
        var (newAudioReference, newMediaReference) = this.AddRelationships(replacementDataPart);
        audioFromFile.Link = newAudioReference.Id;
        p14Media.Embed = newMediaReference.Id;
        var oldDataPart = audioReference.DataPart;
        this.OpenXmlPart.DeleteReferenceRelationship(audioReference.Id);
        this.OpenXmlPart.DeleteReferenceRelationship(mediaReference.Id);
        this.DeleteDataPartIfUnused(oldDataPart);
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

    private A.AudioFromFile AudioFromFile()
    {
        var audioFromFile = pPicture.NonVisualPictureProperties!
            .ApplicationNonVisualDrawingProperties!
            .Descendants<A.AudioFromFile>()
            .SingleOrDefault();
        return audioFromFile ?? throw new SCException("Updating video content is not supported.");
    }

    private (AudioReferenceRelationship AudioReference, MediaReferenceRelationship MediaReference) AddRelationships(
        MediaDataPart mediaDataPart) => this.OpenXmlPart switch
    {
        SlidePart slidePart => (slidePart.AddAudioReferenceRelationship(mediaDataPart), slidePart.AddMediaReferenceRelationship(mediaDataPart)),
        SlideLayoutPart slideLayoutPart => (slideLayoutPart.AddAudioReferenceRelationship(mediaDataPart), slideLayoutPart.AddMediaReferenceRelationship(mediaDataPart)),
        SlideMasterPart slideMasterPart => (slideMasterPart.AddAudioReferenceRelationship(mediaDataPart), slideMasterPart.AddMediaReferenceRelationship(mediaDataPart)),
        _ => throw new SCException("Audio content can only be updated in a presentation part."),
    };

    private MemoryStream BufferedMedia(Stream media)
    {
        if (media.CanSeek)
        {
            media.Position = 0;
        }

        var bufferedMedia = new MemoryStream();
        media.CopyTo(bufferedMedia);
        bufferedMedia.Position = 0;
        return bufferedMedia;
    }

    private MediaDataPart CreateMediaDataPart(MemoryStream media)
    {
        var (contentType, extension) = this.DataPartProperties(media);
        var presentationDocument = (PresentationDocument)this.OpenXmlPart.OpenXmlPackage;
        var mediaDataPart = presentationDocument.CreateMediaDataPart(contentType, extension);
        mediaDataPart.FeedData(media);
        return mediaDataPart;
    }

    private (AudioReferenceRelationship AudioReference, MediaReferenceRelationship MediaReference) ExistingRelationships(
        A.AudioFromFile audioFromFile,
        DocumentFormat.OpenXml.Office2010.PowerPoint.Media p14Media)
    {
        var audioReference = this.OpenXmlPart.DataPartReferenceRelationships
            .OfType<AudioReferenceRelationship>()
            .Single(relationship => relationship.Id == audioFromFile.Link!.Value);
        var mediaReference = this.OpenXmlPart.DataPartReferenceRelationships
            .OfType<MediaReferenceRelationship>()
            .Single(relationship => relationship.Id == p14Media.Embed!.Value);
        return (audioReference, mediaReference);
    }

    private DocumentFormat.OpenXml.Office2010.PowerPoint.Media P14Media() => pPicture.NonVisualPictureProperties!
        .ApplicationNonVisualDrawingProperties!
        .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>()
        .Single();

    private void DeleteDataPartIfUnused(DataPart dataPart)
    {
        if (!dataPart.GetDataPartReferenceRelationships().Any())
        {
            this.OpenXmlPart.OpenXmlPackage.DeletePart(dataPart);
        }
    }

    private (string ContentType, string Extension) DataPartProperties(MemoryStream media)
    {
        var header = new byte[12];
        var read = media.Read(header, 0, header.Length);
        media.Position = 0;
        if (read >= 12 && header[0] == 'R' && header[1] == 'I' && header[2] == 'F' && header[3] == 'F' &&
            header[8] == 'W' && header[9] == 'A' && header[10] == 'V' && header[11] == 'E')
        {
            return ("audio/wav", ".wav");
        }

        if ((read >= 3 && header[0] == 'I' && header[1] == 'D' && header[2] == '3') ||
            (read >= 2 && header[0] == byte.MaxValue && (header[1] & 0xE0) == 0xE0))
        {
            return ("audio/mpeg", ".mp3");
        }

        throw new SCException("Unsupported audio format. Only MP3 and WAVE are supported.");
    }
}
