using System;
using ShapeCrawler;

namespace ShapeCrawler.Tests.Unit
{
    public class TestFile_009Fixture : IDisposable
    {
        public PresentationSc pre009 { get; }

        public TestFile_009Fixture()
        {
            pre009 = PresentationSc.Open(Properties.Resources._009, false);
        }

        public void Dispose()
        {
            pre009.Close();
        }
    }
}