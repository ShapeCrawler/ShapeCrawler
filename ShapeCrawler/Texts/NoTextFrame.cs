using System.Collections.Generic;

namespace ShapeCrawler.Texts
{
    public class NoTextFrame : ITextFrame
    {
        public IList<ParagraphSc> Paragraphs => null;

        public string Text => null;
    }
}