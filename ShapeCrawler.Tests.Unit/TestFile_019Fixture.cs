using System;
using ShapeCrawler.Models;

namespace ShapeCrawler.Tests.Unit
{
    public class TestFile_019Fixture : IDisposable
    {
        public PresentationSc pre019 { get; }

        public TestFile_019Fixture()
        {
            pre019 = new PresentationSc(Properties.Resources._019);
        }

        public void Dispose()
        {
            pre019.Close();
        }
    }
}