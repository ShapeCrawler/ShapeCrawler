using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Factories;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents collection of paragraph text portions.
/// </summary>
public interface IPortionCollection : IEnumerable<IPortion>
{
    /// <summary>
    ///     Gets the number of series items in the collection.
    /// </summary>
    int Count { get; }

    /// <summary>
    ///     Gets the element at the specified index.
    /// </summary>
    IPortion this[int index] { get; }

    /// <summary>
    ///     Adds text portion.
    /// </summary>
    void AddText(string text);
    
    /// <summary>
    /// 	Adds Line Break.
    /// </summary>
    void AddLineBreak();

    /// <summary>
    ///     Removes portion item from collection.
    /// </summary>
    void Remove(IPortion removingPortion);

    /// <summary>
    ///     Removes portion items from collection.
    /// </summary>
    void Remove(IList<IPortion> portions);
}

internal sealed class SCPortionCollection : IPortionCollection
{
    private readonly ResettableLazy<List<IPortion>> portions;
    private readonly A.Paragraph aParagraph;
    private readonly SlideStructure slideStructure;
    private ITextFrameContainer textFrameContainer;

    internal SCPortionCollection(A.Paragraph aParagraph,  SlideStructure slideStructure,ITextFrameContainer textFrameContainer)
    {
        this.aParagraph = aParagraph;
        this.slideStructure = slideStructure;
        this.portions = new ResettableLazy<List<IPortion>>(ParsePortions);
        this.textFrameContainer = textFrameContainer;
    }
    
    public int Count => this.portions.Value.Count;

    public IPortion this[int index] => this.portions.Value[index];

    public void AddText(string text)
    {
        if (text.Contains(Environment.NewLine))
        {
            throw new SCException(
                $"Text can not contain New Line. Use {nameof(IPortionCollection.AddLineBreak)} to add Line Break.");
        }
        
        var lastARunOrABreak = this.aParagraph.LastOrDefault(p => p is A.Run or A.Break);
    
        var lastPortion = (TextPortion?) this.portions.Value.OfType<TextPortion>().First();
        var aTextParent = lastPortion?.AText.Parent ?? new ARunBuilder().Build();

        AddText(ref lastARunOrABreak, aTextParent, text, this.aParagraph);

        this.portions.Reset();
    }

    public void AddLineBreak()
    {
        throw new System.NotImplementedException();
    }

    public void Remove(IPortion removingPortion)
    {
        removingPortion.Remove();

        this.portions.Reset();
    }

    public void Remove(IList<IPortion> removingPortions)
    {
        foreach (var portion in removingPortions)
        {
            this.Remove(portion);
        }
    }

    public IEnumerator<IPortion> GetEnumerator()
    {
        return this.portions.Value.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }
    
    internal void AddNewLine()
    {
        var lastARunOrABreak = this.aParagraph.Last();
        lastARunOrABreak.InsertAfterSelf(new A.Break());
    }

    private static void AddText(ref OpenXmlElement? lastElement, OpenXmlElement aTextParent, string text, A.Paragraph aParagraph)
    {
        var newARun = (A.Run)aTextParent.CloneNode(true);
        newARun.Text!.Text = text;
        if (lastElement == null)
        {
            var apPr = aParagraph.GetFirstChild<A.ParagraphProperties>();
            lastElement = apPr != null ? apPr.InsertAfterSelf(newARun) : aParagraph.InsertAt(newARun, 0);
        }
        else
        {
            lastElement = lastElement.InsertAfterSelf(newARun);
        }
    }

    private List<IPortion> ParsePortions()
    {
        var portions = new List<IPortion>();
        foreach (var paraChild in this.aParagraph.Elements())
        {
            switch (paraChild)
            {
                case A.Run aRun:
                    portions.Add(new TextPortion(aRun, this.slideStructure, this.textFrameContainer));
                    break;
                case A.Field aField:
                {
                    portions.Add(new TextPortion(aField));
                    break;
                }
                case A.Break aBreak:
                    portions.Add(new NewLinePortion(aBreak));
                    break;
            }
        }
        
        return portions;
    }
}