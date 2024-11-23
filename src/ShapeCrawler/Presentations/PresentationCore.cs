using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using ShapeCrawler.Exceptions;

#if NETSTANDARD2_0
using ShapeCrawler.Extensions;
#endif

namespace ShapeCrawler.Presentations;

internal sealed class PresentationCore
{
    private readonly PresentationDocument sdkPresDocument;
    private readonly SlideSize slideSize;

    internal PresentationCore(byte[] bytes)
    {
        var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        this.sdkPresDocument = PresentationDocument.Open(stream, true);
        var sdkMasterParts = this.sdkPresDocument.PresentationPart!.SlideMasterParts;
        this.SlideMasters = new SlideMasterCollection(sdkMasterParts);
        this.Sections = new Sections(this.sdkPresDocument);
        this.Slides = new Slides(this.sdkPresDocument.PresentationPart);
        this.Footer = new Footer(this);
        this.slideSize = new SlideSize(this.sdkPresDocument.PresentationPart!.Presentation.SlideSize!);
    }

    internal PresentationCore(Stream stream)
    {
        stream.Position = 0;
        this.sdkPresDocument = PresentationDocument.Open(stream, true);
        var sdkMasterParts = this.sdkPresDocument.PresentationPart!.SlideMasterParts;
        this.SlideMasters = new SlideMasterCollection(sdkMasterParts);
        this.Sections = new Sections(this.sdkPresDocument);
        this.Slides = new Slides(this.sdkPresDocument.PresentationPart);
        this.Footer = new Footer(this);
        this.slideSize = new SlideSize(this.sdkPresDocument.PresentationPart!.Presentation.SlideSize!);
    }

    internal ISlides Slides { get; }

    internal decimal SlideHeight
    {
        get => this.slideSize.Height();
        set => this.slideSize.UpdateHeight(value);
    }

    internal decimal SlideWidth
    {
        get => this.slideSize.Width();
        set => this.slideSize.UpdateWidth(value);
    }

    internal ISlideMasterCollection SlideMasters { get; }

    internal ISections Sections { get; }

    internal IFooter Footer { get; }

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
        var nonCriticalErrorDesc = new List<string>
        {
                "The element has unexpected child element 'http://schemas.openxmlformats.org/drawingml/2006/chart:showDLblsOverMax'.",
                "The element has invalid child element 'http://schemas.microsoft.com/office/drawing/2017/03/chart:dataDisplayOptions16'. List of possible elements expected: <http://schemas.microsoft.com/office/drawing/2017/03/chart:dispNaAsBlank>.",
                "The 'uri' attribute is not declared.",
                "The 'mod' attribute is not declared.",
                "The 'mod' attribute is not declared.",
                "The element has unexpected child element 'http://schemas.openxmlformats.org/drawingml/2006/main:noFill'.",
                "The element has unexpected child element 'http://schemas.openxmlformats.org/drawingml/2006/main:tc'."
        };
        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(this.sdkPresDocument);
        errors = errors.Where(vr => !nonCriticalErrorDesc.Contains(vr.Description));

#if NETSTANDARD2_0
        errors = errors.DistinctBy(x => new { x.Description, x.Path?.XPath }).ToList();
#else
        errors = errors.DistinctBy(x => new { x.Description, x.Path?.XPath }).ToList();
#endif

        if (errors.Any())
        {
            throw new SCException("Presentation is invalid. See the Errors property for details.");
        }
    }
}