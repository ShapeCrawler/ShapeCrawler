using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents collection of paragraph text portions.
    /// </summary>
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    [SuppressMessage("ReSharper", "PossibleMultipleEnumeration")]
    [SuppressMessage("ReSharper", "SuggestVarOrType_BuiltInTypes")]
    [SuppressMessage("ReSharper", "SuggestVarOrType_Elsewhere")]
    internal class PortionCollection : EditableCollection<IPortion>, IPortionCollection
    {
        public override void Remove(IPortion portion)
        {
            if (portion == null || !CollectionItems.Contains(portion))
            {
                return;
            }

            CollectionItems.Remove(portion);

            ((Portion) portion).AText.Parent.Remove(); // removes from DOM
        }

        public void Remove(IList<IPortion> removingPortions)
        {
            foreach (var portion in removingPortions)
            {
                CollectionItems.Remove(portion);
                ((Portion) portion).AText.Parent.Remove();
            }
        }

        #region Internal Methods

        internal PortionCollection(List<IPortion> portions)
        {
            CollectionItems = portions;
        }

        /// <summary>
        ///     Gets collection of paragraph portions. Returns <c>NULL</c> if paragraph is empty.
        /// </summary>
        internal static PortionCollection Create(A.Paragraph aParagraph, SCParagraph paragraph)
        {
            IEnumerable<A.Run> aRuns = aParagraph.Elements<A.Run>();
            if (aRuns.Any())
            {
                var runPortions = new List<IPortion>(aRuns.Count());
                foreach (A.Run aRun in aRuns)
                {
                    runPortions.Add(new Portion(aRun.Text, paragraph));
                }

                return new PortionCollection(runPortions);
            }

            A.Field aField = aParagraph.GetFirstChild<A.Field>();
            if (aField != null)
            {
                A.Text aText = aParagraph.GetFirstChild<A.Field>().GetFirstChild<A.Text>();
                var aFieldPortions = new List<IPortion>(new[] {new Portion(aText, paragraph)});
                return new PortionCollection(aFieldPortions);
            }

            return null;
        }

        #endregion Internal Methods
    }
}