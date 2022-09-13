using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Collections;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable SuggestVarOrType_BuiltInTypes
namespace ShapeCrawler.Texts
{
    [SuppressMessage("ReSharper", "PossibleMultipleEnumeration")]
    internal class ParagraphCollection : IParagraphCollection
    {
        private readonly ResettableLazy<List<SCParagraph>> paragraphs;
        private readonly SCTextBox textBox;

        internal ParagraphCollection(SCTextBox textBox)
        {
            this.textBox = textBox;
            this.paragraphs = new ResettableLazy<List<SCParagraph>>(this.GetParagraphs);
        }

        #region Public Properties

        public int Count => this.paragraphs.Value.Count;

        public IParagraph this[int index] => this.paragraphs.Value[index];

        public IEnumerator<IParagraph> GetEnumerator()
        {
            return this.paragraphs.Value.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        #endregion Public Properties
        
        public IParagraph Add()
        {
            var lastAParagraph = this.paragraphs.Value.Last().AParagraph;
            var newAParagraph = (A.Paragraph)lastAParagraph.CloneNode(true);
            lastAParagraph.InsertAfterSelf(newAParagraph);

            var newParagraph = new SCParagraph(newAParagraph, this.textBox)
            {
                Text = string.Empty
            };

            this.paragraphs.Reset();

            return newParagraph;
        }

        public void Remove(IEnumerable<IParagraph> removeParagraphs)
        {
            foreach (var paragraph in removeParagraphs.Cast<SCParagraph>())
            {
                paragraph.AParagraph.Remove();
                paragraph.IsRemoved = true;
            }

            this.paragraphs.Reset();
        }

        private List<SCParagraph> GetParagraphs()
        {
            if (this.textBox.APTextBody == null)
            {
                return new List<SCParagraph>(0);
            }

            var aParagraphs = this.textBox.APTextBody.Elements<A.Paragraph>();
            var nonEmptyAPara = aParagraphs.Where(p => p.Elements<A.Run>().Any() || p.Elements<A.Field>().Any());
            return nonEmptyAPara.Select(aParagraph => new SCParagraph(aParagraph, this.textBox)).ToList();
        }
    }
}