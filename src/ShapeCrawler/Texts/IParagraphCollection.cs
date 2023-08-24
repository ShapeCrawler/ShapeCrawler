using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a collection of paragraphs.
/// </summary>
public interface IParagraphCollection : IReadOnlyList<IParagraph>
{
    /// <summary>
    ///     Adds a new paragraph in collection.
    /// </summary>
    IParagraph Add();

    /// <summary>
    ///     Removes specified paragraphs from collection.
    /// </summary>
    void Remove(IEnumerable<IParagraph> removeParagraphs);
}

internal sealed class Paragraphs : IParagraphCollection
{
    private readonly ResetableLazy<List<Paragraph>> paragraphs;

    internal Paragraphs()
    {
        this.paragraphs = new ResetableLazy<List<Paragraph>>(this.ParseParagraphs);
    }

    #region Public Properties

    public int Count => this.paragraphs.Value.Count;

    public IParagraph this[int index] => this.paragraphs.Value[index];

    public IEnumerator<IParagraph> GetEnumerator()
    {
        return this.paragraphs.Value.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    #endregion Public Properties

    public IParagraph Add()
    {
        var lastAParagraph = this.paragraphs.Value.Last().AParagraph;
        var newAParagraph = (A.Paragraph)lastAParagraph.CloneNode(true);
        newAParagraph.ParagraphProperties ??= new A.ParagraphProperties();
        lastAParagraph.InsertAfterSelf(newAParagraph);

        var newParagraph = new Paragraph(newAParagraph)
        {
            Text = string.Empty
        };

        this.paragraphs.Reset();

        return newParagraph;
    }

    public void Remove(IEnumerable<IParagraph> removeParagraphs)
    {
        foreach (var paragraph in removeParagraphs.Cast<Paragraph>())
        {
            paragraph.AParagraph.Remove();
            paragraph.IsRemoved = true;
        }

        this.paragraphs.Reset();
    }

    private List<Paragraph> ParseParagraphs()
    {
        if (this.textFrame.TextBodyElement == null)
        {
            return new List<Paragraph>(0);
        }

        var paraList = new List<Paragraph>();
        foreach (var aPara in this.textFrame.TextBodyElement.Elements<A.Paragraph>())
        {
            var para = new Paragraph(aPara, this.textFrame, this.slideStructure, this.textFrameContainer);
            para.TextChanged += this.textFrame.OnParagraphTextChanged;
            paraList.Add(para);
        }

        return paraList;
    }
}