using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Factories;
using ShapeCrawler.Shared;
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
    ///     Adds portion item to collection.
    /// </summary>
    void Add(string newPortionText);

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
    private readonly ResettableLazy<List<SCPortion>> portions;
    private readonly A.Paragraph aParagraph;
    private readonly SCParagraph parentParagraph;

    internal SCPortionCollection(A.Paragraph aParagraph, SCParagraph paragraph)
    {
        this.aParagraph = aParagraph;
        this.parentParagraph = paragraph;
        this.portions = new ResettableLazy<List<SCPortion>>(() => GetPortions(aParagraph, paragraph));
    }
    
    public int Count => this.portions.Value.Count;

    public IPortion this[int index] => this.portions.Value[index];

    public void Add(string newPortionText)
    {
        var lastARunOrABreak = this.aParagraph.LastOrDefault(p => p is A.Run or A.Break);
    
        var lastPortion = this.portions.Value.LastOrDefault();
        var aTextParent = lastPortion?.AText.Parent ?? new ARunBuilder().Build();

        AddText(ref lastARunOrABreak, aTextParent, newPortionText, this.aParagraph);

        this.portions.Reset();
    }
    
    public void Remove(IPortion removingPortion)
    {
        var removingInnerPortion = (SCPortion)removingPortion;

        removingInnerPortion.AText.Parent!.Remove(); // remove parent <a:r>
        removingInnerPortion.IsRemoved = true;

        this.portions.Reset();
    }

    public void Remove(IList<IPortion> removingPortions)
    {
        foreach (SCPortion portion in removingPortions.Cast<SCPortion>())
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

    private static List<SCPortion> GetPortions(A.Paragraph aParagraph, SCParagraph paragraph)
    {
        var aRuns = aParagraph.Elements<A.Run>();
        if (aRuns.Any())
        {
            var runPortions = new List<SCPortion>(aRuns.Count());
            foreach (var aRun in aRuns)
            {
                runPortions.Add(new SCPortion(aRun.Text!, paragraph));
            }

            return runPortions;
        }

        var aField = aParagraph.GetFirstChild<A.Field>();
        if (aField != null)
        {
            var aText = aField.GetFirstChild<A.Text>();
            var newPortion = new SCPortion(aText!, paragraph, aField);
            var fieldPortions = new List<SCPortion>(new[] { newPortion });
            return fieldPortions;
        }

        return new List<SCPortion>();
    }
}