using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using ShapeCrawler.Exceptions;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

internal sealed record Presentation
{
    private readonly PresentationDocument sdkPresDocument;
    private readonly SlideSize slideSize;

    internal Presentation(byte[] bytes)
        : this(new MemoryStream(bytes))
    {
    }

    internal Presentation(Stream stream)
    {
        stream.Position = 0;
        this.sdkPresDocument = PresentationDocument.Open(stream, true);
        var sdkMasterParts = this.sdkPresDocument.PresentationPart!.SlideMasterParts;
        this.SlideMasters = new SlideMasterCollection(sdkMasterParts);
        this.Sections = new Sections(this.sdkPresDocument);
        this.Slides = new Slides(this.sdkPresDocument.PresentationPart!.SlideParts);
        this.HeaderAndFooter = new HeaderAndFooter(this);
        this.slideSize = new SlideSize(this.sdkPresDocument.PresentationPart!.Presentation.SlideSize!);
    }

    internal ISlides Slides { get; }

    internal int SlideHeight
    {
        get => this.slideSize.Height();
        set => this.slideSize.UpdateHeight(value);
    }

    internal int SlideWidth
    {
        get => this.slideSize.Width();
        set => this.slideSize.UpdateWidth(value);
    }

    internal ISlideMasterCollection SlideMasters { get; }

    internal ISections Sections { get; }

    internal IHeaderAndFooter HeaderAndFooter { get; }

    internal void CopyTo(string path)
    {
        var cloned = this.sdkPresDocument.Clone(path);
        cloned.Dispose();
    }

    internal void CopyTo(Stream stream) => this.sdkPresDocument.Clone(stream);

    internal byte[] AsByteArray()
    {
        var stream = new MemoryStream();
        this.sdkPresDocument.Clone(stream);

        return stream.ToArray();
    }

    internal void Validate()
    {
        var nonCritical = new List<ValidationError>
        {
            new(
                "The element has unexpected child element 'http://schemas.openxmlformats.org/drawingml/2006/chart:showDLblsOverMax'.",
                "/c:chartSpace[1]/c:chart[1]"),
            new("/c:chartSpace[1]/c:chart[1]/c:extLst[1]/c:ext[1]", "/c:chartSpace[1]/c:chart[1]"),
            new(
                "The element has invalid child element 'http://schemas.microsoft.com/office/drawing/2017/03/chart:dataDisplayOptions16'. List of possible elements expected: <http://schemas.microsoft.com/office/drawing/2017/03/chart:dispNaAsBlank>.",
                "/c:chartSpace[1]/c:chart[1]/c:extLst[1]/c:ext[1]"),
            new(
                "The 'uri' attribute is not declared.",
                "/c:chartSpace[1]/c:chart[1]/c:extLst[1]/c:ext[1]"),
            new(
                "The 'mod' attribute is not declared.",
                "/p:sldLayout[1]/p:extLst[1]"),
            new(
                "The 'mod' attribute is not declared.",
                "/p:sldMaster[1]/p:extLst[1]"),
            new(
                "The element has unexpected child element 'http://schemas.openxmlformats.org/drawingml/2006/main:pPr'.",
                "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[1]/p:txBody[1]/a:p[1]")
        };

        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        var errors = validator.Validate(this.sdkPresDocument);

        var removing = new List<ValidationErrorInfo>();
        foreach (var error in errors)
        {
            if (nonCritical.Any(x => x.Description == error.Description && x.Path == error.Path?.XPath))
            {
                removing.Add(error);
            }
        }

        errors = errors.Except(removing).DistinctBy(x => new { x.Description, x.Path?.XPath }).ToList();

        if (errors.Any())
        {
            throw new SCException("Presentation is invalid. See the Errors property for details.");
        }
    }
}