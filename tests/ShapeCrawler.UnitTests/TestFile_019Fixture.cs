using System;
using ShapeCrawler.Models;
using SlideDotNet.Models;

namespace ShapeCrawler.UnitTests
{
    public class TestFile_019Fixture : IDisposable
    {
        public Presentation pre019 { get; }

        public TestFile_019Fixture()
        {
            pre019 = new Presentation(Properties.Resources._019);
        }

        public void Dispose()
        {
            pre019.Close();
        }
    }
}