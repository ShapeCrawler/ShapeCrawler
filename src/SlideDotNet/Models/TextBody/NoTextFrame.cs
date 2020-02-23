using System.Collections.Generic;
using SlideDotNet.Exceptions;

namespace SlideDotNet.Models.TextBody
{
    public class NoTextFrame : ITextFrame
    {
        public IList<Paragraph> Paragraphs => throw new SlideDotNetException(ExceptionMessages.NoTextFrame);

        public string Text => throw new SlideDotNetException(ExceptionMessages.NoTextFrame);
    }
}