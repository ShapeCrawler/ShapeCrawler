using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ShapeCrawler.Settings;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Collections
{
    /// <summary>
    /// Represents collection of paragraph text portions.
    /// </summary>
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    [SuppressMessage("ReSharper", "PossibleMultipleEnumeration")]
    [SuppressMessage("ReSharper", "SuggestVarOrType_BuiltInTypes")]
    [SuppressMessage("ReSharper", "SuggestVarOrType_Elsewhere")]
    public class PortionCollection : EditableCollection<Portion>
    {
        #region Internal Methods

        internal PortionCollection(List<Portion> portions)
        {
            CollectionItems = portions;
        }

        internal static PortionCollection Create(A.Paragraph aParagraph, ShapeContext spContext, ParagraphSc paragraph)
        {
            IEnumerable<A.Run> aRuns = aParagraph.Elements<A.Run>();
            if (aRuns.Any())
            {
                var portions = new List<Portion>(aRuns.Count());
                foreach (A.Run aRun in aRuns)
                {
                    var newPortion = new Portion(aRun.Text, paragraph, spContext);
                    portions.Add(newPortion);
                }

                return new PortionCollection(portions);
            }
            else
            {
                A.Text aText = aParagraph.GetFirstChild<A.Field>().GetFirstChild<A.Text>();
                var newPortion = new Portion(aText, paragraph, spContext);
                var portions = new List<Portion>(new[] { newPortion });

                return new PortionCollection(portions);
            }
        }

        #endregion Internal Methods

        public override void Remove(Portion portion)
        {
            if (!CollectionItems.Contains(portion))
            {
                return;
            }
            CollectionItems.Remove(portion);

            portion.AText.Parent.Remove(); // removes from DOM
        }

        public void Remove(IList<Portion> removingPortions)
        {
            foreach (var portion in removingPortions)
            {
                CollectionItems.Remove(portion);
                portion.AText.Parent.Remove();
            }
        }
    }
}