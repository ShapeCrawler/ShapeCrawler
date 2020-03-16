using SlideDotNet.Collections;
using SlideDotNet.Enums;
using SlideDotNet.Exceptions;

namespace SlideDotNet.Models.SlideComponents.Chart
{
    public class NoChart : IChart
    {
        public ChartType Type => throw new SlideDotNetException(ExceptionMessages.NoChart);

        public string Title => throw new SlideDotNetException(ExceptionMessages.NoChart);

        public bool HasTitle => throw new SlideDotNetException(ExceptionMessages.NoChart);

        public SeriesCollection SeriesCollection => throw new SlideDotNetException(ExceptionMessages.NoChart);

        public CategoryCollection Categories => throw new SlideDotNetException(ExceptionMessages.NoChart);

        public bool HasCategories => throw new SlideDotNetException(ExceptionMessages.NoChart);
    }
}