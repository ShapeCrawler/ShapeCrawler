using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable SuggestBaseTypeForParameter

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     <inheritdoc cref="IPortionCollection"/>
    /// </summary>
    internal class PortionCollection : IPortionCollection
    {
        private readonly ResettableLazy<List<SCPortion>> portions;

        /// <summary>
        ///     Initializes a new instance of the <see cref="PortionCollection"/> class.
        /// </summary>
        public PortionCollection(A.Paragraph aParagraph, SCParagraph paragraph)
        {
            this.portions = new ResettableLazy<List<SCPortion>>(() => GetPortions(aParagraph, paragraph));
        }

        /// <inheritdoc/>
        public int Count => this.portions.Value.Count;

        /// <inheritdoc/>
        public IPortion this[int index] => this.portions.Value[index];

        /// <inheritdoc/>
        public void Remove(IPortion removingPortion)
        {
            SCPortion removingInnerPortion = (SCPortion)removingPortion;

            removingInnerPortion.SDKAText.Parent.Remove(); // remove parent <a:r>
            removingInnerPortion.IsRemoved = true;

            this.portions.Reset();
        }

        /// <inheritdoc/>
        public void Remove(IList<IPortion> removingPortions)
        {
            foreach (SCPortion portion in removingPortions.Cast<SCPortion>())
            {
                this.Remove(portion);
            }
        }

        /// <inheritdoc/>
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
                    runPortions.Add(new SCPortion(aRun.Text, paragraph));
                }

                return runPortions;
            }

            var aField = aParagraph.GetFirstChild<A.Field>();
            if (aField != null)
            {
                var aText = aField.GetFirstChild<A.Text>();
                var aFieldPortions = new List<SCPortion>(new[] { new SCPortion(aText, paragraph) });
                return aFieldPortions;
            }

            return new List<SCPortion>();
        }
    }
}