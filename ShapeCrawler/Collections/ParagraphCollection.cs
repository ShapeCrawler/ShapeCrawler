using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Collections;
using ShapeCrawler.Settings;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable SuggestVarOrType_BuiltInTypes

namespace ShapeCrawler.Texts
{
    public class ParagraphCollection : LibraryCollection<ParagraphSc>
    {
        private readonly ShapeContext _spContext;
        private readonly TextBoxSc _parentTextBox;

        #region Constructors

        internal ParagraphCollection() : base(new List<ParagraphSc>())
        {
        }

        internal ParagraphCollection(IEnumerable<ParagraphSc> paragraphItems, ShapeContext spContext, TextBoxSc parentTextBox) 
            : base(paragraphItems)
        {
            _spContext = spContext;
            _parentTextBox = parentTextBox;
        }

        #endregion Constructors

        internal static ParagraphCollection Create(OpenXmlCompositeElement textBodyCompositeElement, ShapeContext spContext, TextBoxSc parentTextBox)
        {
            // Parse non-empty paragraphs
            IEnumerable<A.Paragraph> nonEmptyAParagraphs = textBodyCompositeElement.Elements<A.Paragraph>().Where(p => p.Descendants<A.Text>().Any());

            var paragraphs = new List<ParagraphSc>(nonEmptyAParagraphs.Count());
            paragraphs.AddRange(nonEmptyAParagraphs.Select(aParagraph => new ParagraphSc(spContext, aParagraph, parentTextBox)));

            return new ParagraphCollection(paragraphs, spContext, parentTextBox);
        }

        #region Public Methods

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

        /// <summary>
        /// Adds a new paragraph in collection.
        /// </summary>
        /// <returns>Added <see cref="ParagraphSc"/> instance.</returns>
        public ParagraphSc Add()
        {
            // Create a new paragraph from the last paragraph and insert at the end
            A.Paragraph lastAParagraph = CollectionItems.Last().AParagraph;
            A.Paragraph newAParagraph = (A.Paragraph) lastAParagraph.CloneNode(true);
            lastAParagraph.InsertAfterSelf(newAParagraph);

            ParagraphSc newParagraph = new ParagraphSc(_spContext, newAParagraph, _parentTextBox);
            newParagraph.Text = string.Empty;
            
            CollectionItems.Add(newParagraph);

            return newParagraph;
        }

        #endregion Public Methods
    }
}