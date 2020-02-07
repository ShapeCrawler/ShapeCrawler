using SlideXML.Enums;
using SlideXML.Exceptions;

namespace SlideXML.Models.SlideComponents.Chart
{
    /// <inheritdoc cref="IChart"/>
    public class NoChart : IChart
    {
        public ChartType Type => throw new SlideXMLException(ExceptionMessages.NoChart);

        public string Title => throw new SlideXMLException(ExceptionMessages.NoChart);

        public bool HasTitle => throw new SlideXMLException(ExceptionMessages.NoChart);
    }
}