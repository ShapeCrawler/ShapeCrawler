using System;
using ShapeCrawler;

namespace ShapeCrawler.Tests.Unit
{
    public class TestFile_009Fixture : IDisposable
    {
        public PresentationSc pre009 { get; }

        public TestFile_009Fixture()
        {
            pre009 = new PresentationSc(Properties.Resources._009);
        }

        public void Dispose()
        {
            pre009.Close();
        }
    }
}