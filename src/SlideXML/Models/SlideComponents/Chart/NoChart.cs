using SlideXML.Enums;
using SlideXML.Exceptions;

namespace SlideXML.Models.SlideComponents.Chart
{
    public class NoChart : IChart
    {
        public ChartType Type => throw new SlideXMLException(ExceptionMessages.NoChart);

        public string Title => throw new SlideXMLException(ExceptionMessages.NoChart);
    }
}