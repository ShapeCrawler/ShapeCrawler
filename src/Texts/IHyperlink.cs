using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a hyperlink.
/// </summary>
public interface IHyperlink
{
    /// <summary>
    ///     Gets the file or web address, or null for an absent or internal slide link.
    /// </summary>
    string? File { get; }

    /// <summary>
    ///     Gets the linked slide's current one-based number, or null for an absent or external link.
    /// </summary>
    int? SlideNumber { get; }

    /// <summary>
    ///     Sets the click link to a nonempty file or web address, replacing the previous click link.
    /// </summary>
    void AddFile(string file);

    /// <summary>
    ///     Sets the click link to a slide in the current presentation, replacing the previous click link.
    ///     Internal links can be added on regular slides.
    /// </summary>
    /// <param name="slide">The one-based target slide number.</param>
    void AddSlideNumber(int slide);

    /// <summary>
    ///     Removes the click hyperlink without changing text, formatting, mouse-over links, or target slides.
    ///     Does nothing if no click hyperlink is assigned.
    /// </summary>
    void Remove();
}

internal class Hyperlink(OpenXmlCompositeElement properties) : IHyperlink
{
    private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public int? SlideNumber => this.GetSlideNumberOrNull();

    public string? File => this.GetFileOrNull();

    private OpenXmlPartRootElement Root => properties.Ancestors<OpenXmlPartRootElement>().First();

    private OpenXmlPart Part => this.Root.OpenXmlPart!;

    public void AddSlideNumber(int slide)
    {
        var presentation = ((PresentationDocument)this.Part.OpenXmlPackage).PresentationPart!;
        var slideIds = presentation.Presentation!.SlideIdList!.Elements<P.SlideId>().ToArray();
        if (slide < 1 || slide > slideIds.Length)
        {
            throw new ArgumentOutOfRangeException(nameof(slide));
        }

        if (this.Part is not SlidePart currentSlidePart)
        {
            throw new SCException("Internal slide hyperlinks can only be added on regular slides.");
        }

        var targetSlidePart = (SlidePart)presentation.GetPartById(slideIds[slide - 1].RelationshipId!);
        var addedPart = currentSlidePart.AddPart(targetSlidePart);
        var hyperlink = this.Click();
        var oldId = hyperlink.Id?.Value;
        hyperlink.Id = currentSlidePart.GetIdOfPart(addedPart);
        hyperlink.Action = "ppaction://hlinksldjump";
        this.RemoveUnusedExternalRelationship(oldId);
    }

    public void AddFile(string file)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(file);
        var uri = new Uri(file, UriKind.RelativeOrAbsolute);
        var relationship = this.Part.HyperlinkRelationships.FirstOrDefault(link => link.IsExternal && link.Uri.OriginalString == uri.OriginalString)
                           ?? this.Part.AddHyperlinkRelationship(uri, true);
        var hyperlink = this.Click();
        var oldId = hyperlink.Id?.Value;
        hyperlink.Id = relationship.Id;
        hyperlink.Action = null;
        this.RemoveUnusedExternalRelationship(oldId);
    }

    public void Remove()
    {
        var hyperlink = properties.GetFirstChild<A.HyperlinkOnClick>();
        if (hyperlink is null)
        {
            return;
        }

        var oldId = hyperlink.Id?.Value;
        hyperlink.Remove();
        this.RemoveUnusedExternalRelationship(oldId);
    }

    private A.HyperlinkOnClick Click()
    {
        var hyperlink = properties.GetFirstChild<A.HyperlinkOnClick>();
        if (hyperlink is null)
        {
            hyperlink = new A.HyperlinkOnClick();
            properties.AddChild(hyperlink, true);
        }

        return hyperlink;
    }

    private void RemoveUnusedExternalRelationship(string? id)
    {
        var relationship = this.Part.HyperlinkRelationships.FirstOrDefault(link => link.Id == id);
        if (relationship is null)
        {
            return;
        }

        var used = this.Root.Descendants().Any(element => element.GetAttributes().Any(attribute =>
            attribute.NamespaceUri == RelationshipNamespace && attribute.Value == id));
        if (!used)
        {
            this.Part.DeleteReferenceRelationship(relationship);
        }
    }

    private string? GetFileOrNull()
    {
        var hyperlink = properties.GetFirstChild<A.HyperlinkOnClick>();
        if (hyperlink is null || hyperlink.Action?.Value == "ppaction://hlinksldjump")
        {
            return null;
        }

        return this.Part.HyperlinkRelationships.FirstOrDefault(link => link.Id == hyperlink.Id?.Value)?.Uri.ToString();
    }

    private int? GetSlideNumberOrNull()
    {
        var hyperlink = properties.GetFirstChild<A.HyperlinkOnClick>();
        if (hyperlink == null || hyperlink.Action?.Value != "ppaction://hlinksldjump" ||
            string.IsNullOrEmpty(hyperlink.Id))
        {
            return null;
        }

        var presentation = ((PresentationDocument)this.Part.OpenXmlPackage).PresentationPart!;
        if (!this.Part.TryGetPartById(hyperlink.Id!, out var target) || target is not SlidePart targetSlidePart)
        {
            return null;
        }

        var slideIdList = presentation.Presentation!.SlideIdList!.ChildElements.OfType<P.SlideId>();
        var targetSlideRelationshipId = presentation.GetIdOfPart(targetSlidePart);

        var index = 0;
        foreach (var slideId in slideIdList)
        {
            index++;
            if (slideId.RelationshipId == targetSlideRelationshipId)
            {
                return index;
            }
        }

        return null;
    }
}
