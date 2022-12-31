using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable SuggestBaseTypeForParameter
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Collections;

internal sealed class PortionCollection : IPortionCollection
{
    private readonly ResettableLazy<List<SCPortion>> portions;

    internal PortionCollection(A.Paragraph aParagraph, SCParagraph paragraph)
    {
        this.portions = new ResettableLazy<List<SCPortion>>(() => GetPortions(aParagraph, paragraph));
    }

    public int Count => this.portions.Value.Count;

    public IPortion this[int index] => this.portions.Value[index];

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
            var fieldPortions = new List<SCPortion>(new[] { newPortion});
            return fieldPortions;
        }

        return new List<SCPortion>();
    }
}