using System;
using ShapeCrawler.Models;

namespace ShapeCrawler.UnitTests.Helpers
{
    public class ReadOnlyTestPresentations : IDisposable
    {
        public Presentation Pre001 { get; }
        public Presentation Pre002 { get; }
        public Presentation Pre003 { get; }
        public Presentation Pre006 { get; }
        public Presentation Pre008 { get; }
        public Presentation Pre009 { get; }
        public Presentation Pre013 { get; }
        public Presentation Pre021 { get; }

        public ReadOnlyTestPresentations()
        {
            Pre001 = Presentation.Open(Properties.Resources._001, false);
            Pre002 = Presentation.Open(Properties.Resources._002, false);
            Pre003 = Presentation.Open(Properties.Resources._003, false);
            Pre006 = Presentation.Open(Properties.Resources._006_1_slides, false);
            Pre008 = Presentation.Open(Properties.Resources._008, false);
            Pre009 = Presentation.Open(Properties.Resources._009, false);
            Pre013 = Presentation.Open(Properties.Resources._013, false);
            Pre021 = Presentation.Open(Properties.Resources._021, false);
        }

        public void Dispose()
        {

        }
    }
}