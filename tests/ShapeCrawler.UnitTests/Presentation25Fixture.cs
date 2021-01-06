using System;
using ShapeCrawler.Models;

namespace ShapeCrawler.Tests.Unit
{
    /// <summary>
    /// Represents a class whose single instance is shared for tests.
    /// </summary>
    public class Presentation25Fixture : IDisposable
    {
        public Presentation Presentation { get; }

        public Presentation25Fixture()
        {
            Presentation = new Presentation(Properties.Resources._025);
        }

        public void Dispose()
        {
            Presentation.Close();
        }
    }
}