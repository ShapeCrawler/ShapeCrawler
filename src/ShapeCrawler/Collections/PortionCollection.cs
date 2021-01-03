using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Models.TextBody;
using ShapeCrawler.NoLogic;
using ShapeCrawler.Settings;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Collections
{
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    public class PortionCollection : EditableCollection<Portion>
    {
        // TODO: Delete one of collection which duplicates collection of Portions.
        // _portionToText and CollectionItems both store Portion items.
        private readonly Dictionary<Portion, A.Text> _portionToText;

        public PortionCollection(List<Portion> portions, Dictionary<Portion, A.Text> portionToText)
        {
            CollectionItems = portions;
            _portionToText = portionToText;
        }

        public override void Remove(Portion portion)
        {
            if (!CollectionItems.Contains(portion))
            {
                return;
            }
            CollectionItems.Remove(portion);

            _portionToText[portion].Parent.Remove(); // removes from DOM
        }

        public void RemoveRange(IList<Portion> removingPortions)
        {
            foreach (var portion in removingPortions)
            {
                CollectionItems.Remove(portion);
                var aText = _portionToText[portion];
                aText.Parent.Remove();
            }
        }

        public static PortionCollection Create(A.Paragraph aParagraph, IShapeContext spContext, int innerPrLvl, Paragraph paragraph)
        {
            var regularTextRuns = aParagraph.Elements<A.Run>();
            if (regularTextRuns.Any())
            {
                var portions = new List<Portion>(regularTextRuns.Count());
                var portionToText = new Dictionary<Portion, A.Text>(regularTextRuns.Count());
                foreach (var run in regularTextRuns)
                {
                    int fh = FontHeightFromRun(run, spContext, innerPrLvl);
                    A.Text aText = run.Text;
                    var newPortion = new Portion(aText, fh, paragraph);
                    portions.Add(newPortion);
                    portionToText.Add(newPortion, aText);
                }

                return new PortionCollection(portions, portionToText);
            }
            else
            {
                A.Text aText = aParagraph.GetFirstChild<A.Field>().GetFirstChild<A.Text>();
                int fh = FontHeightFromOther(spContext, innerPrLvl);
                var newPortion = new Portion(aText, fh, paragraph);
                var portions = new List<Portion>(new[] {newPortion});
                var portionToText = new Dictionary<Portion, A.Text> {{newPortion, aText}};

                return new PortionCollection(portions, portionToText);
            }
        }

        private static int FontHeightFromRun(A.Run run, IShapeContext spContext, int innerPrLvl)
        {
            var runFs = run.RunProperties?.FontSize;
            var resultFh = runFs == null ? FontHeightFromOther(spContext, innerPrLvl) : runFs.Value;

            return resultFh;
        }

        private static int FontHeightFromOther(IShapeContext spContext, int innerPrLvl)
        {
            // if element is placeholder, tries to get from placeholder data
            var xmlElement = spContext.SdkElement;
            if (xmlElement.IsPlaceholder())
            {
                var prFontHeight =
                    spContext.PlaceholderFontService.TryGetFontHeight((OpenXmlCompositeElement)xmlElement,
                        innerPrLvl);
                if (prFontHeight != null)
                {
                    return (int)prFontHeight;
                }
            }

            if (spContext.presentationData.LlvFontHeights.ContainsKey(innerPrLvl))
            {
                return spContext.presentationData.LlvFontHeights[innerPrLvl];
            }

            var exist = spContext.TryGetFontHeight(innerPrLvl, out int fh);
            if (exist)
            {
                return fh;
            }

            return FormatConstants.DefaultFontHeight;
        }
    }
}