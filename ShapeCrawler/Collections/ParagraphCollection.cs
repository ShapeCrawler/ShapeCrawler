using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Collections;
using ShapeCrawler.Settings;
using A = DocumentFormat.OpenXml.Drawing;
// ReSharper disable CheckNamespace
// ReSharper disable SuggestVarOrType_BuiltInTypes

namespace ShapeCrawler.Texts
{
    public class ParagraphCollection : LibraryCollection<ParagraphSc>
    {
        public ParagraphCollection(IEnumerable<ParagraphSc> items) : base(items)
        {
        }

        public static ParagraphCollection Parse(OpenXmlCompositeElement compositeElement, ShapeContext spContext, TextSc text)
        {
            // Parse non-empty paragraphs
            IEnumerable<A.Paragraph> nonEmptyAParagraphs = compositeElement.Elements<A.Paragraph>().Where(p => p.Descendants<A.Text>().Any());

            var paragraphs = new List<ParagraphSc>(nonEmptyAParagraphs.Count());
            foreach (A.Paragraph aParagraph in nonEmptyAParagraphs)
            {
                paragraphs.Add(new ParagraphSc(spContext, aParagraph, text));
            }

            return new ParagraphCollection(paragraphs);
        }

        public void RemoveRange(int index, int removeCount)
        {
            // Remove from outer
            for (int removeIdx = index; removeIdx <= removeCount; removeIdx++)
            {
                CollectionItems[removeIdx].Remove();
            }

            // Remove from collection
            CollectionItems.RemoveRange(index, removeCount);
        }
    }
}