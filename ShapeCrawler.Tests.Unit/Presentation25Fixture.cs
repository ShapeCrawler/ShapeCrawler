using System;
using ShapeCrawler.Models;

namespace ShapeCrawler.Tests.Unit
{
    /// <summary>
    /// Represents a class whose single instance is shared for tests.
    /// </summary>
    public class Presentation25Fixture : IDisposable
    {
        public PresentationEx PresentationEx { get; }

        public Presentation25Fixture()
        {
            PresentationEx = new PresentationEx(Properties.Resources._025);
        }

        public void Dispose()
        {
            PresentationEx.Close();
        }
    }
}