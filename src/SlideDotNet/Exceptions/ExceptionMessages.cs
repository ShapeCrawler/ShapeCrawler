using SlideDotNet.Models.SlideComponents.Chart;

namespace SlideDotNet.Exceptions
{
    /// <summary>
    /// Contains constant error messages.
    /// </summary>
    public static class ExceptionMessages
    {
        public static string NoTextFrame = "Element has not a text frame.";

        public static string NoChart = "Element has not a chart.";

        public static string NoPicture = "Element has not a picture.";

        public static string NoTable = "Element has not a table.";

        public static string NoOleObject = "Element has not a OLE object.";

        public static string NotPlaceholder = "Element is not a placeholder";

        public static string NotTitle = "Chart has not a title.";

        /// <summary>
        /// Returns message string with placeholder.
        /// </summary>
        public static string PresentationIsLarge = "The size of presentation more than {0} bytes.";

        public static string SlidesMuchMore = "The number of slides is more allowed {0}.";

        public static string ChartCanNotHaveCategory = $"#0 can not have category. You can check chart type via {nameof(ChartEx.Type)} property.";
    }
}
