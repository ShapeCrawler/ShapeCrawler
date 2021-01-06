using System;
using System.Collections.Generic;
using ShapeCrawler.Exceptions;

namespace ShapeCrawler.Models.TextBody
{
    public class NoTextFrame : ITextFrame
    {
        public IList<Paragraph> Paragraphs => null;

        public string Text => null;
    }
}