using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ShapeCrawler.AutoShapes;
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
    public class PortionCollection : EditableCollection<Portion>
    {
        public override void Remove(Portion row)
        {
            if (!CollectionItems.Contains(row))
            {
                return;
            }

            CollectionItems.Remove(row);

            row.AText.Parent.Remove(); // removes from DOM
        }

        public void Remove(IList<Portion> removingPortions)
        {
            foreach (var portion in removingPortions)
            {
                CollectionItems.Remove(portion);
                portion.AText.Parent.Remove();
            }
        }

        #region Internal Methods

        internal PortionCollection(List<Portion> portions)
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
                var runPortions = new List<Portion>(aRuns.Count());
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
                var aFieldPortions = new List<Portion>(new[] {new Portion(aText, paragraph)});
                return new PortionCollection(aFieldPortions);
            }

            return null;
        }

        #endregion Internal Methods
    }
}