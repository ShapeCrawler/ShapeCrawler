using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

/// <summary>
///     Represents a hyperlink.
/// </summary>
public interface IHyperlink
{
    /// <summary>
    ///     Gets or sets linked slide number. Returns <see langword="null"/> if hyperlink is not a slide link.
    /// </summary>
    int? SlideNumber { get; set; }

    /// <summary>
    ///     Gets or sets URL address. Returns <see langword="null"/> if hyperlink is not a URL link.
    /// </summary>
    string? Url { get; set; }
}

internal class Hyperlink(RunProperties aRunProperties) : IHyperlink
{
    public int? SlideNumber
    {
        get => GetSlideNumber();
        set => SetSlideNumber(value);
    }

    public string? Url { get; set; }

    private void SetSlideNumber(int? slideNumber)
    {
        if (slideNumber is null)
        {
            throw new SCException("The specified slide number is null.");
        }

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
            .ElementAtOrDefault(slideNumber.Value - 1)!;

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

    private void SetLink(string? address)
    {
        if (address is null)
        {
            throw new SCException("The specified link address is null.");
        }

        var hyperlink = aRunProperties.GetFirstChild<A.HyperlinkOnClick>();
        if (hyperlink == null)
        {
            hyperlink = new A.HyperlinkOnClick();
            aRunProperties.Append(hyperlink);
        }
            var partRoot = aRunProperties.Ancestors<OpenXmlPartRootElement>().First();
            var uri = new Uri(address, UriKind.RelativeOrAbsolute);
            var addedHyperlinkRelationship = partRoot.OpenXmlPart!.AddHyperlinkRelationship(uri, true);
            hyperlink.Id = addedHyperlinkRelationship.Id;
            hyperlink.Action = null;
    }

    private void AddSlideLink(string url, A.HyperlinkOnClick hyperlink)
    {
        // Handle inner slide hyperlink
        var slideNumber = int.Parse(url.Substring(8));
        var parentXmlPart = aRunProperties.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var presentation = ((PresentationDocument)parentXmlPart.OpenXmlPackage).PresentationPart!;
        var slideId = presentation.Presentation.SlideIdList!.ChildElements
            .OfType<P.SlideId>()
            .ElementAtOrDefault(slideNumber - 1)!;

        // Get the target slide part
        var targetSlidePart = (SlidePart)presentation.GetPartById(slideId.RelationshipId!);

        // Add relationship from current slide to target slide
        var currentSlidePart = (SlidePart)parentXmlPart;

        // Add or reuse relationship to target slide
        var relationship = currentSlidePart.GetIdOfPart(targetSlidePart);

        hyperlink.Id = relationship;
        hyperlink.Action = "ppaction://hlinksldjump";
    }


    private int? GetSlideNumber()
    {
        var hyperlink = aRunProperties.GetFirstChild<A.HyperlinkOnClick>();
        if (hyperlink == null || hyperlink.Action?.Value != "ppaction://hlinksldjump" || string.IsNullOrEmpty(hyperlink.Id))
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