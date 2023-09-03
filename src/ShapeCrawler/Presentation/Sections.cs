using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler;

internal sealed class Sections : ISections
{
    private readonly PresentationDocument sdkPresDocument;
    
    internal Sections (PresentationDocument sdkPresDocument)
    {
        this.sdkPresDocument = sdkPresDocument;
    }

    public int Count => this.SectionList().Count;

    private List<Section> SectionList()
    {
        var p14SectionList = this.sdkPresDocument.PresentationPart!.Presentation.PresentationExtensionList
            ?.Descendants<P14.SectionList>().FirstOrDefault();
        return p14SectionList == null 
            ? new List<Section>(0) 
            : p14SectionList.OfType<P14.Section>().Select(p14Section => new Section(this.sdkPresDocument, p14Section)).ToList();
    }

    public ISection this[int index] => this.SectionList()[index];

    public IEnumerator<ISection> GetEnumerator()
    {
        return this.SectionList().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    public void Remove(ISection removingSection)
    {
        if (removingSection is not IRemoveable removeable)
        {
            throw new SCException("Section cannot be removed.");
        }
        
        var total = this.Count;
        removeable.Remove();

        if (total == 1)
        {
            this.sdkPresDocument.PresentationPart!.Presentation.PresentationExtensionList
                ?.Descendants<P14.SectionList>().First()
                .Remove();
        }
    }

    public ISection GetByName(string sectionName)
    {
        return this.SectionList().First(section => section.Name == sectionName);
    }
}