using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Services.Factories;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents collection of paragraph text portions.
/// </summary>
public interface IParagraphPortions : IEnumerable<IParagraphPortion>
{
    /// <summary>
    ///     Gets the number of series items in the collection.
    /// </summary>
    int Count { get; }

    /// <summary>
    ///     Gets the element at the specified index.
    /// </summary>
    IParagraphPortion this[int index] { get; }

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
    void Remove(IParagraphPortion removingPortion);

    /// <summary>
    ///     Removes portion items from collection.
    /// </summary>
    void Remove(IList<IParagraphPortion> portions);
}

internal sealed class ParagraphPortions : IParagraphPortions
{
    private readonly TypedOpenXmlPart sdkTypedOpenXmlPart;
    private readonly A.Paragraph aParagraph;

    internal ParagraphPortions(TypedOpenXmlPart sdkTypedOpenXmlPart, A.Paragraph aParagraph)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.aParagraph = aParagraph;
    }

    public int Count => this.Portions().Count;
    public IParagraphPortion this[int index] => this.Portions()[index];

    public void AddText(string text)
    {
        if (text.Contains(Environment.NewLine))
        {
            throw new SCException(
                $"Text can not contain New Line. Use {nameof(IParagraphPortions.AddLineBreak)} to add Line Break.");
        }

        var lastARunOrABreak = this.aParagraph.LastOrDefault(p => p is A.Run or A.Break);

        var textPortions = this.Portions().OfType<TextParagraphPortion>();
        var lastPortion = textPortions.Any() ? textPortions.Last() : null;
        var aTextParent = lastPortion?.AText.Parent ?? new ARunBuilder().Build();

        AddText(ref lastARunOrABreak, aTextParent, text, this.aParagraph);
    }

    public void AddLineBreak()
    {
        throw new System.NotImplementedException();
    }

    public void Remove(IParagraphPortion removingPortion)
    {
        removingPortion.Remove();
    }

    public void Remove(IList<IParagraphPortion> removingPortions)
    {
        foreach (var portion in removingPortions)
        {
            this.Remove(portion);
        }
    }

    public IEnumerator<IParagraphPortion> GetEnumerator() => this.Portions().GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    internal void AddNewLine()
    {
        var lastARunOrABreak = this.aParagraph.Last();
        lastARunOrABreak.InsertAfterSelf(new A.Break());
    }

    private static void AddText(ref OpenXmlElement? lastElement, OpenXmlElement aTextParent, string text,
        A.Paragraph aParagraph)
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

    private List<IParagraphPortion> Portions()
    {
        var portions = new List<IParagraphPortion>();
        foreach (var aParagraphElement in this.aParagraph.Elements())
        {
            switch (aParagraphElement)
            {
                case A.Run aRun:
                    var runPortion = new TextParagraphPortion(this.sdkTypedOpenXmlPart, aRun);
                    portions.Add(runPortion);
                    break;
                case A.Field aField:
                {
                    var fieldPortion = new Field(this.sdkTypedOpenXmlPart, aField);
                    portions.Add(fieldPortion);
                    break;
                }
                case A.Break aBreak:
                    var lineBreak = new ParagraphLineBreak(aBreak);
                    portions.Add(lineBreak);
                    break;
            }
        }

        return portions;
    }
}