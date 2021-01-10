using System;
using ShapeCrawler.Models;

namespace ShapeCrawler.Tests.Unit.Helpers
{
    public class PptxFixture : IDisposable
    {
        public PresentationEx Pre001 { get; }
        public PresentationEx Pre002 { get; }
        public PresentationEx Pre003 { get; }
        public PresentationEx Pre006 { get; }
        public PresentationEx Pre008 { get; }
        public PresentationEx Pre009 { get; }
        public PresentationEx Pre013 { get; }
        public PresentationEx Pre015 { get; }
        public PresentationEx Pre016 { get; }
        public PresentationEx Pre017 { get; }
        public PresentationEx Pre018 { get; }
        public PresentationEx Pre020 { get; }
        public PresentationEx Pre021 { get; }
        public PresentationEx Pre024 { get; }

        public PptxFixture()
        {
            Pre001 = PresentationEx.Open(Properties.Resources._001, false);
            Pre002 = PresentationEx.Open(Properties.Resources._002, false);
            Pre003 = PresentationEx.Open(Properties.Resources._003, false);
            Pre006 = PresentationEx.Open(Properties.Resources._006_1_slides, false);
            Pre008 = PresentationEx.Open(Properties.Resources._008, false);
            Pre009 = PresentationEx.Open(Properties.Resources._009, false);
            Pre013 = PresentationEx.Open(Properties.Resources._013, false);
            Pre015 = PresentationEx.Open(Properties.Resources._015, false);
            Pre016 = PresentationEx.Open(Properties.Resources._016, false);
            Pre017 = PresentationEx.Open(Properties.Resources._017, false);
            Pre018 = PresentationEx.Open(Properties.Resources._018, false);
            Pre020 = PresentationEx.Open(Properties.Resources._020, false);
            Pre021 = PresentationEx.Open(Properties.Resources._021, false);
            Pre024 = PresentationEx.Open(Properties.Resources._024, false);
        }

        public void Dispose()
        {

        }
    }
}