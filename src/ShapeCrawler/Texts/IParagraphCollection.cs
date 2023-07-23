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

internal sealed class ParagraphCollection : IParagraphCollection
{
    private readonly ResetAbleLazy<List<SCParagraph>> paragraphs;
    private readonly SCTextFrame textFrame;
    private readonly SlideStructure slideStructure;
    private readonly ITextFrameContainer textFrameContainer;

    internal ParagraphCollection(SCTextFrame textFrame, SlideStructure slideStructure, ITextFrameContainer textFrameContainer)
    {
        this.textFrame = textFrame;
        this.slideStructure = slideStructure;
        this.paragraphs = new ResetAbleLazy<List<SCParagraph>>(this.GetParagraphs);
        this.textFrameContainer = textFrameContainer;
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

        var newParagraph = new SCParagraph(newAParagraph, this.textFrame, this.slideStructure,this.textFrameContainer)
        {
            Text = string.Empty
        };

        this.paragraphs.Reset();

        return newParagraph;
    }

    public void Remove(IEnumerable<IParagraph> removeParagraphs)
    {
        foreach (var paragraph in removeParagraphs.Cast<SCParagraph>())
        {
            paragraph.AParagraph.Remove();
            paragraph.IsRemoved = true;
        }

        this.paragraphs.Reset();
    }

    private List<SCParagraph> GetParagraphs()
    {
        if (this.textFrame.TextBodyElement == null)
        {
            return new List<SCParagraph>(0);
        }

        var paraList = new List<SCParagraph>();
        foreach (var aPara in this.textFrame.TextBodyElement.Elements<A.Paragraph>())
        {
            var para = new SCParagraph(aPara, this.textFrame, this.slideStructure, this.textFrameContainer);
            para.TextChanged += this.textFrame.OnParagraphTextChanged;
            paraList.Add(para);
        }

        return paraList;
    }
}