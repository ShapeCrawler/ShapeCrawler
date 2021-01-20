namespace ShapeCrawler.Tests.Unit.Helpers
{
    public class PresentationFixture
    {
        public PresentationSc Pre001 { get; }
        public PresentationSc Pre002 { get; }
        public PresentationSc Pre003 { get; }
        public PresentationSc Pre006 { get; }
        public PresentationSc Pre008 { get; }
        public PresentationSc Pre009 { get; }
        public PresentationSc Pre013 { get; }
        public PresentationSc Pre014 { get; }
        public PresentationSc Pre015 { get; }
        public PresentationSc Pre016 { get; }
        public PresentationSc Pre017 { get; }
        public PresentationSc Pre018 { get; }
        public PresentationSc Pre019 { get; }
        public PresentationSc Pre020 { get; }
        public PresentationSc Pre021 { get; }
        public PresentationSc Pre024 { get; }
        public PresentationSc Pre025 { get; }

        public PresentationFixture()
        {
            Pre001 = PresentationSc.Open(Properties.Resources._001, false);
            Pre002 = PresentationSc.Open(Properties.Resources._002, false);
            Pre003 = PresentationSc.Open(Properties.Resources._003, false);
            Pre006 = PresentationSc.Open(Properties.Resources._006_1_slides, false);
            Pre008 = PresentationSc.Open(Properties.Resources._008, false);
            Pre009 = PresentationSc.Open(Properties.Resources._009, false);
            Pre013 = PresentationSc.Open(Properties.Resources._013, false);
            Pre014 = PresentationSc.Open(Properties.Resources._014, false);
            Pre015 = PresentationSc.Open(Properties.Resources._015, false);
            Pre016 = PresentationSc.Open(Properties.Resources._016, false);
            Pre017 = PresentationSc.Open(Properties.Resources._017, false);
            Pre018 = PresentationSc.Open(Properties.Resources._018, false);
            Pre019 = PresentationSc.Open(Properties.Resources._019, false);
            Pre020 = PresentationSc.Open(Properties.Resources._020, false);
            Pre021 = PresentationSc.Open(Properties.Resources._021, false);
            Pre024 = PresentationSc.Open(Properties.Resources._024, false);
            Pre025 = PresentationSc.Open(Properties.Resources._025, false);
        }
    }
}