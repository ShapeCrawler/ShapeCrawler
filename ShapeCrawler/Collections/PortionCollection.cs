using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Models.TextShape;
using ShapeCrawler.Settings;
using ShapeCrawler.Statics;
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

        internal static PortionCollection Create(A.Paragraph aParagraph, ShapeContext spContext, int innerPrLvl, ParagraphEx paragraphEx)
        {
            IEnumerable<A.Run> aRuns = aParagraph.Elements<A.Run>();
            if (aRuns.Any())
            {
                var portions = new List<Portion>(aRuns.Count());
                foreach (A.Run aRun in aRuns)
                {
                    A.Text aText = aRun.Text;
                    var newPortion = new Portion(aText, paragraphEx, spContext, innerPrLvl);
                    portions.Add(newPortion);
                }

                return new PortionCollection(portions);
            }
            else
            {
                A.Text aText = aParagraph.GetFirstChild<A.Field>().GetFirstChild<A.Text>();
                var newPortion = new Portion(aText, paragraphEx, spContext, innerPrLvl);
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

        public void RemoveRange(IList<Portion> removingPortions)
        {
            foreach (var portion in removingPortions)
            {
                CollectionItems.Remove(portion);
                portion.AText.Parent.Remove();
            }
        }
    }
}