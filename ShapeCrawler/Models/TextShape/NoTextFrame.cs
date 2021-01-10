using System.Collections.Generic;

namespace ShapeCrawler.Models.TextShape
{
    public class NoTextFrame : ITextFrame
    {
        public IList<ParagraphEx> Paragraphs => null;

        public string Text => null;
    }
}