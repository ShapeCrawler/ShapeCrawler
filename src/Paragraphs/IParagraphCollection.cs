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

    /// <summary>
    ///     Adds a new paragraph at the specified index of the collection.
    /// </summary>
    void Add(string content, int index);
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
        var aParagraphs = textBody.Elements<A.Paragraph>().ToList();
        if (index < 0 || index > aParagraphs.Count)
        {
            throw new System.ArgumentOutOfRangeException(nameof(index));
        }

        if (index == aParagraphs.Count)
        {
            this.Add();
            this.ParagraphsCore().Last().Text = content;
        }
        else
        {
            var refParagraph = aParagraphs[index];

            // Preserve paragraph properties
            var pPr = refParagraph.GetFirstChild<A.ParagraphProperties>()
                             ?.CloneNode(true) as A.ParagraphProperties;

            // Clone and clear children
            var newAParagraph = (A.Paragraph)refParagraph.CloneNode(true);
            newAParagraph.RemoveAllChildren();
            if (pPr != null)
            {
                newAParagraph.Append(pPr);
            }
            else
            {
                newAParagraph.ParagraphProperties = new A.ParagraphProperties();
            }

            // Create new run with content
            var firstRun = refParagraph.Elements<A.Run>().FirstOrDefault();
            A.Run newRun;
            if (firstRun != null)
            {
                var newRunPr = firstRun.RunProperties?.CloneNode(true) as A.RunProperties
                               ?? new A.RunProperties();
                var aText = new A.Text { Text = content };
                newRun = new A.Run(newRunPr, aText);
            }
            else
            {
                var newRunPr = new A.RunProperties { Language = "en-US", Dirty = false };
                var aText = new A.Text { Text = content };
                newRun = new A.Run(newRunPr, aText);
            }

            newAParagraph.Append(newRun);
            refParagraph.InsertBeforeSelf(newAParagraph);
        }
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