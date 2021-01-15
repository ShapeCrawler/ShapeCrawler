using System;
using ShapeCrawler;

namespace ShapeCrawler.Tests.Unit
{
    /// <summary>
    /// Represents a class whose single instance is shared for tests.
    /// </summary>
    public class Presentation25Fixture : IDisposable
    {
        public PresentationSc PresentationEx { get; }

        public Presentation25Fixture()
        {
            PresentationEx = new PresentationSc(Properties.Resources._025);
        }

        public void Dispose()
        {
            PresentationEx.Close();
        }
    }
}