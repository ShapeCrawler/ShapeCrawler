using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning disable IDE0130

/// <summary>
///     Represents a collection of paragraphs.
/// </summary>
public interface IParagraphCollection : IReadOnlyList<IParagraph>
{
    /// <summary>
    ///     Adds a new paragraph at the end of the collection.
    /// </summary>
    void Add();

#if DEBUG
    /// <summary>
    ///     Adds a new paragraph at the specified index of the collection.
    /// </summary>
    void Add(string content, int index);
#endif
}

internal readonly struct ParagraphCollection(OpenXmlElement textBody): IParagraphCollection
{
    public int Count => this.ParagraphsCore().Count;

    public IParagraph this[int index] => this.ParagraphsCore()[index];

    public IEnumerator<IParagraph> GetEnumerator() => this.ParagraphsCore().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public void Add()
    {
        var lastAParagraph = textBody.Elements<A.Paragraph>().Last();
        var newAParagraph = (A.Paragraph)lastAParagraph.CloneNode(true);
        newAParagraph.ParagraphProperties ??= new A.ParagraphProperties();
        lastAParagraph.InsertAfterSelf(newAParagraph);
    }

    public void Add(string content, int index)
    {
        throw new System.NotImplementedException();
    }

    private List<Paragraph> ParagraphsCore()
    {
        var aParagraphs = textBody.Elements<A.Paragraph>().ToList();
        if (!aParagraphs.Any())
        {
            return [];
        }

        var paraList = new List<Paragraph>();
        foreach (var aPara in aParagraphs)
        {
            var para = new Paragraph(aPara);
            paraList.Add(para);
        }

        return paraList;
    }
}