using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Exceptions;

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
    }
    
    private void SetLink(string? address)
    {
        if (address is null)
        {
            throw new SCException("The specified link address is null.");
        }
        
        var runProperties = this.AText.PreviousSibling<A.RunProperties>() ?? new A.RunProperties();

        var hyperlink = runProperties.GetFirstChild<A.HyperlinkOnClick>();
        if (hyperlink == null)
        {
            hyperlink = new A.HyperlinkOnClick();
            runProperties.Append(hyperlink);
        }

        if (address.StartsWith("slide://"))
        {
            AddSlideLink(address, hyperlink);
        }
        else
        {
            var uri = new Uri(address, UriKind.RelativeOrAbsolute);
            var addedHyperlinkRelationship = this.sdkTypedOpenXmlPart.AddHyperlinkRelationship(uri, true);
            hyperlink.Id = addedHyperlinkRelationship.Id;
            hyperlink.Action = null;
        }
    }
    
    private void AddSlideLink(string url, A.HyperlinkOnClick hyperlink)
    {
        // Handle inner slide hyperlink
        var slideNumber = int.Parse(url.Substring(8));
        var presentation = ((PresentationDocument)this.sdkTypedOpenXmlPart.OpenXmlPackage).PresentationPart!;
        var slideId = presentation.Presentation.SlideIdList!.ChildElements
            .OfType<P.SlideId>()
            .ElementAtOrDefault(slideNumber - 1)!;

        // Get the target slide part
        var targetSlidePart = (SlidePart)presentation.GetPartById(slideId.RelationshipId!);

        // Add relationship from current slide to target slide
        var currentSlidePart = (SlidePart)this.sdkTypedOpenXmlPart;

        // Add or reuse relationship to target slide
        var relationship = currentSlidePart.GetIdOfPart(targetSlidePart);

        hyperlink.Id = relationship;
        hyperlink.Action = "ppaction://hlinksldjump";
    }


    private int? GetSlideNumber()
    {
        throw new System.NotImplementedException();
    }
}