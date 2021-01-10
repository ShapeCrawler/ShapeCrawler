using System;
using ShapeCrawler.Models;

namespace ShapeCrawler.Tests.Unit
{
    public class TestFile_009Fixture : IDisposable
    {
        public PresentationEx pre009 { get; }

        public TestFile_009Fixture()
        {
            pre009 = new PresentationEx(Properties.Resources._009);
        }

        public void Dispose()
        {
            pre009.Close();
        }
    }
}