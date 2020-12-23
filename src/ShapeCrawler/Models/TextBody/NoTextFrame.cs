using System;
using System.Collections.Generic;
using ShapeCrawler.Exceptions;

namespace ShapeCrawler.Models.TextBody
{
    public class NoTextFrame : ITextFrame
    {
        public IList<Paragraph> Paragraphs => throw new NotSupportedException(ExceptionMessages.NoTextFrame);

        public string Text => throw new NotSupportedException(ExceptionMessages.NoTextFrame);
    }
}