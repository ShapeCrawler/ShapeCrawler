using System;
using ShapeCrawler.Models;

namespace ShapeCrawler.UnitTests
{
    /// <summary>
    /// Represents a class whose single instance is shared for tests.
    /// </summary>
    public class Presentation25Fixture : IDisposable
    {
        public PresentationEx Presentation { get; }

        public Presentation25Fixture()
        {
            Presentation = new PresentationEx(Properties.Resources._025);
        }

        public void Dispose()
        {
            Presentation.Close();
        }
    }
}