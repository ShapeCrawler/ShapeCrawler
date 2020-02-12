using SlideXML.Enums;
using SlideXML.Exceptions;

namespace SlideXML.Models.SlideComponents.Chart
{
    /// <inheritdoc cref="IChart"/>
    public class NoChart : IChart
    {
        public ChartType Type => throw new SlideXmlException(ExceptionMessages.NoChart);

        public string Title => throw new SlideXmlException(ExceptionMessages.NoChart);

        public bool HasTitle => throw new SlideXmlException(ExceptionMessages.NoChart);
    }
}