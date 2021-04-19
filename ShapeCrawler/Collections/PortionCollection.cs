using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     <inheritdoc cref="IPortionCollection"/>
    /// </summary>
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    [SuppressMessage("ReSharper", "PossibleMultipleEnumeration")]
    [SuppressMessage("ReSharper", "SuggestVarOrType_BuiltInTypes")]
    [SuppressMessage("ReSharper", "SuggestVarOrType_Elsewhere")]
    internal class PortionCollection : IPortionCollection
    {
        private readonly ResettableLazy<List<Portion>> portions;

        /// <summary>
        ///     Initializes a new instance of the <see cref="PortionCollection"/> class.
        /// </summary>
        public PortionCollection(A.Paragraph aParagraph, SCParagraph paragraph)
        {
            this.portions = new ResettableLazy<List<Portion>>(() => this.GetPortions(aParagraph, paragraph));
        }

        /// <summary>
        ///     Gets the number of paragraph portions.
        /// </summary>
        public int Count => this.portions.Value.Count;

        public IPortion this[int index] => this.portions.Value[index];

        public void Remove(IPortion removingPortion)
        {
            if (removingPortion == null || !this.portions.Value.Contains(removingPortion))
            {
                return;
            }

            ((Portion)removingPortion).AText.Parent.Remove();

            this.portions.Reset();
        }

        public void Remove(IList<IPortion> removingPortions)
        {
            foreach (var portion in removingPortions)
            {
                ((Portion)portion).AText.Parent.Remove();
            }

            this.portions.Reset();
        }

        /// <summary>
        ///     Gets an enumerator that iterates through the paragraph portion collection.
        /// </summary>
        public IEnumerator<IPortion> GetEnumerator()
        {
            return this.portions.Value.GetEnumerator();
        }

        public List<Portion> GetPortions (A.Paragraph aParagraph, SCParagraph paragraph)
        {
            IEnumerable<A.Run> aRuns = aParagraph.Elements<A.Run>();
            if (aRuns.Any())
            {
                var runPortions = new List<Portion>(aRuns.Count());
                foreach (A.Run aRun in aRuns)
                {
                    runPortions.Add(new Portion(aRun.Text, paragraph));
                }

                return runPortions;
            }

            A.Field aField = aParagraph.GetFirstChild<A.Field>();
            if (aField != null)
            {
                A.Text aText = aParagraph.GetFirstChild<A.Field>().GetFirstChild<A.Text>();
                var aFieldPortions = new List<Portion>(new[] {new Portion(aText, paragraph)});
                return aFieldPortions;
            }

            return new List<Portion>();
        }

        

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }
}