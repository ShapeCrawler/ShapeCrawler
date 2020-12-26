using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Models.Settings;
using ShapeCrawler.Models.TextBody;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Collections
{
    public class PortionCollection : EditableCollection<Portion>
    {
        public PortionCollection(List<Portion> portions)
        {
            CollectionItems = portions;
        }

        public override void Remove(Portion item)
        {
            CollectionItems.Remove(item);
        }

        public static PortionCollection Create(A.Paragraph aParagraph, IShapeContext spContext, int innerPrLvl, Paragraph paragraph)
        {
            var regularTextRuns = aParagraph.Elements<A.Run>();
            if (regularTextRuns.Any())
            {
                var portions = new List<Portion>(regularTextRuns.Count());
                foreach (var run in regularTextRuns)
                {
                    var fh = FontHeightFromRun(run, spContext, innerPrLvl);
                    portions.Add(new Portion(run.Text, fh, paragraph));
                }

                return new PortionCollection(portions);
            }
            else
            {
                var textField = aParagraph.GetFirstChild<A.Field>()
                    .GetFirstChild<A.Text>();
                var fh = FontHeightFromOther(spContext, innerPrLvl);
                var portions = new List<Portion>(1)
                {
                    new Portion(textField, fh, paragraph)
                };

                return new PortionCollection(portions);
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

            if (spContext.PreSettings.LlvFontHeights.ContainsKey(innerPrLvl))
            {
                return spContext.PreSettings.LlvFontHeights[innerPrLvl];
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