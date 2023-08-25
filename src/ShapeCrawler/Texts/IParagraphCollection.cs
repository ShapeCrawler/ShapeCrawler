using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
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
    void Add();

    /// <summary>
    ///     Removes specified paragraphs from collection.
    /// </summary>
    void Remove(IEnumerable<IParagraph> removeParagraphs);
}

internal sealed class Paragraphs : IParagraphCollection
{
    private readonly IEnumerable<A.Paragraph> aParagraphs;
    private readonly ResetableLazy<List<SlideParagraph>> paragraphs;
    private readonly SlidePart sdkSlidePart;

    internal Paragraphs(SlidePart sdkSlidePart, IEnumerable<A.Paragraph> aParagraphs)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.aParagraphs = aParagraphs;
        this.paragraphs = new ResetableLazy<List<SlideParagraph>>(this.ParseParagraphs);
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

    public void Add()
    {
        var lastAParagraph = this.paragraphs.Value.Last().AParagraph;
        var newAParagraph = (A.Paragraph)lastAParagraph.CloneNode(true);
        newAParagraph.ParagraphProperties ??= new A.ParagraphProperties();
        lastAParagraph.InsertAfterSelf(newAParagraph);
        
        this.paragraphs.Reset();
    }

    public void Remove(IEnumerable<IParagraph> removeParagraphs)
    {
        foreach (var paragraph in removeParagraphs.Cast<SlideParagraph>())
        {
            paragraph.AParagraph.Remove();
            paragraph.IsRemoved = true;
        }

        this.paragraphs.Reset();
    }

    private List<SlideParagraph> ParseParagraphs()
    {
        if (!this.aParagraphs.Any())
        {
            return new List<SlideParagraph>(0);
        }

        var paraList = new List<SlideParagraph>();
        foreach (var aPara in this.aParagraphs)
        {
            var para = new SlideParagraph(this.sdkSlidePart, aPara);
            paraList.Add(para);
        }

        return paraList;
    }
}