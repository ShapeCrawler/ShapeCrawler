using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning disable IDE0130

/// <summary>
///     Represents a collection of paragraphs.
/// </summary>
public interface IParagraphs : IReadOnlyList<IParagraph>
{
    /// <summary>
    ///     Adds a new paragraph in collection.
    /// </summary>
    void Add();
}

internal readonly struct Paragraphs(OpenXmlPart openXmlPart, OpenXmlElement sdkTextBody) : IParagraphs
{
    public int Count => this.ParagraphsCore().Count;
    
    public IParagraph this[int index] => this.ParagraphsCore()[index];
    
    public IEnumerator<IParagraph> GetEnumerator() => this.ParagraphsCore().GetEnumerator();
    
    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public void Add()
    {
        var lastAParagraph = sdkTextBody.Elements<A.Paragraph>().Last();
        var newAParagraph = (A.Paragraph)lastAParagraph.CloneNode(true);
        newAParagraph.ParagraphProperties ??= new A.ParagraphProperties();
        lastAParagraph.InsertAfterSelf(newAParagraph);
    }

    private List<Paragraph> ParagraphsCore()
    {
        var aParagraphs = sdkTextBody.Elements<A.Paragraph>().ToList();
        if (!aParagraphs.Any())
        {
            return [];
        }

        var paraList = new List<Paragraph>();
        foreach (var aPara in aParagraphs)
        {
            var para = new Paragraph(openXmlPart, aPara);
            paraList.Add(para);
        }

        return paraList;
    }
}