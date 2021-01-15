using System.Collections.Generic;

namespace ShapeCrawler.Models.TextShape
{
    public class NoTextFrame : ITextFrame
    {
        public IList<Paragraph> Paragraphs => null;

        public string Text => null;
    }
}