using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
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
    ///     Gets the address to the existing File or Web Page.
    /// </summary>
    string? File { get; }

    /// <summary>
    ///     Gets the number of the linked slide.
    /// </summary>
    int? SlideNumber { get; }

    /// <summary>
    ///     Adds the address to the existing File or Web Page.
    /// </summary>
    void AddFile(string file);

    /// <summary>
    ///     Adds the link to the slide.
    /// </summary>
    /// <param name="slide">The number of the linking slide.</param>
    void AddSlideNumber(int slide);
}

internal class Hyperlink(RunProperties aRunProperties): IHyperlink
{
    public int? SlideNumber => this.GetSlideNumberOrNull();

    public string? File => this.GetFileOrNull();
    
    public void AddSlideNumber(int slide)
    {
        var hyperlink = aRunProperties.GetFirstChild<A.HyperlinkOnClick>();
        if (hyperlink == null)
        {
            hyperlink = new A.HyperlinkOnClick();
            aRunProperties.Append(hyperlink);
        }

        var parentXmlPart = aRunProperties.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var presentation = ((PresentationDocument)parentXmlPart.OpenXmlPackage).PresentationPart!;
        var slideId = presentation.Presentation.SlideIdList!.ChildElements
            .OfType<P.SlideId>()
            .ElementAtOrDefault(slide - 1) !;

        // Get the target slide part
        var targetSlidePart = (SlidePart)presentation.GetPartById(slideId.RelationshipId!);

        // Add relationship from current slide to target slide
        var currentSlidePart = (SlidePart)parentXmlPart;

        // Add or reuse relationship to target slide
        var addedPart = currentSlidePart.AddPart(targetSlidePart);
        var relationship = currentSlidePart.GetIdOfPart(addedPart);

        hyperlink.Id = relationship;
        hyperlink.Action = "ppaction://hlinksldjump";
    }

    public void AddFile(string file)
    {
        var hyperlink = aRunProperties.GetFirstChild<A.HyperlinkOnClick>();
        if (hyperlink == null)
        {
            hyperlink = new A.HyperlinkOnClick();
            aRunProperties.Append(hyperlink);
        }

        var partRoot = aRunProperties.Ancestors<OpenXmlPartRootElement>().First();
        var uri = new Uri(file, UriKind.RelativeOrAbsolute);
        var addedHyperlinkRelationship = partRoot.OpenXmlPart!.AddHyperlinkRelationship(uri, true);
        hyperlink.Id = addedHyperlinkRelationship.Id;
        hyperlink.Action = null;
    }

    private string? GetFileOrNull()
    {
        var hyperlink = aRunProperties.GetFirstChild<A.HyperlinkOnClick>();
        if (hyperlink == null)
        {
            return null;
        }

        var parentXmlPart = aRunProperties.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var hyperlinkRelationship = (HyperlinkRelationship)parentXmlPart.GetReferenceRelationship(hyperlink.Id!);

        return hyperlinkRelationship.Uri.ToString();
    }

    private int? GetSlideNumberOrNull()
    {
        var hyperlink = aRunProperties.GetFirstChild<A.HyperlinkOnClick>();
        if (hyperlink == null || hyperlink.Action?.Value != "ppaction://hlinksldjump" ||
            string.IsNullOrEmpty(hyperlink.Id))
        {
            return null;
        }

        var parentXmlPart = aRunProperties.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var presentation = ((PresentationDocument)parentXmlPart.OpenXmlPackage).PresentationPart!;
        var currentSlidePart = (SlidePart)parentXmlPart;
        var targetSlidePart = (SlidePart)currentSlidePart.GetPartById(hyperlink.Id!);

        var slideIdList = presentation.Presentation.SlideIdList!.ChildElements.OfType<P.SlideId>();
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