using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Shapes;

internal sealed class ShapeRelationships(OpenXmlPart sourcePart, OpenXmlPart targetPart)
{
    private readonly Dictionary<string, string> relationshipIds = [];
    private readonly Dictionary<DataPart, MediaDataPart> mediaParts = [];

    internal void CopyTo(OpenXmlElement shape)
    {
        if (sourcePart == targetPart)
        {
            return;
        }

        foreach (var element in shape.Descendants().Prepend(shape))
        {
            foreach (var attribute in element.GetAttributes())
            {
                if (attribute.NamespaceUri != "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    || string.IsNullOrEmpty(attribute.Value))
                {
                    continue;
                }

                if (!this.relationshipIds.TryGetValue(attribute.Value, out var targetId))
                {
                    targetId = this.CopyRelationship(attribute.Value);
                    this.relationshipIds.Add(attribute.Value, targetId);
                }

                element.SetAttribute(new OpenXmlAttribute(
                    attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, targetId));
            }
        }
    }

    private string CopyRelationship(string sourceId)
    {
        if (sourcePart.TryGetPartById(sourceId, out var part))
        {
            var copiedPart = targetPart.AddPart(part!);
            return targetPart.GetIdOfPart(copiedPart);
        }

        var dataRelationship = sourcePart.DataPartReferenceRelationships.FirstOrDefault(r => r.Id == sourceId);
        if (dataRelationship is not null)
        {
            return this.CopyDataRelationship(dataRelationship);
        }

        var hyperlink = sourcePart.HyperlinkRelationships.FirstOrDefault(r => r.Id == sourceId);
        if (hyperlink is not null)
        {
            return targetPart.AddHyperlinkRelationship(hyperlink.Uri, hyperlink.IsExternal).Id;
        }

        var external = sourcePart.ExternalRelationships.Single(r => r.Id == sourceId);
        return targetPart.AddExternalRelationship(external.RelationshipType, external.Uri).Id;
    }

    private string CopyDataRelationship(DataPartReferenceRelationship relationship)
    {
        var sourceMedia = relationship.DataPart;
        if (!this.mediaParts.TryGetValue(sourceMedia, out var targetMedia))
        {
            if (sourcePart.OpenXmlPackage == targetPart.OpenXmlPackage)
            {
                targetMedia = (MediaDataPart)sourceMedia;
            }
            else
            {
                targetMedia = targetPart.OpenXmlPackage.CreateMediaDataPart(
                    sourceMedia.ContentType, Path.GetExtension(sourceMedia.Uri.OriginalString));
                using var stream = sourceMedia.GetStream();
                targetMedia.FeedData(stream);
            }

            this.mediaParts.Add(sourceMedia, targetMedia);
        }

        var slidePart = (SlidePart)targetPart;
        return relationship switch
        {
            AudioReferenceRelationship => slidePart.AddAudioReferenceRelationship(targetMedia).Id,
            VideoReferenceRelationship => slidePart.AddVideoReferenceRelationship(targetMedia).Id,
            MediaReferenceRelationship => slidePart.AddMediaReferenceRelationship(targetMedia).Id,
            _ => throw new SCException("Unsupported media relationship type.")
        };
    }
}
