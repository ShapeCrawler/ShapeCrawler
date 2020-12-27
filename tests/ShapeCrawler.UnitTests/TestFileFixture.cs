using System;
using ShapeCrawler.Models;

namespace ShapeCrawler.UnitTests
{
    public class TestFileFixture : IDisposable
    {
        public Presentation Pre002 { get; }
        public Presentation Pre003 { get; }
        public Presentation Pre013 { get; }
        public Presentation Pre021 { get; }

        public TestFileFixture()
        {
            Pre002 = Presentation.Open(Properties.Resources._002, false);
            Pre003 = Presentation.Open(Properties.Resources._003, false);
            Pre013 = Presentation.Open(Properties.Resources._013, false);
            Pre021 = Presentation.Open(Properties.Resources._021, false);
        }

        public void Dispose()
        {

        }
    }
}