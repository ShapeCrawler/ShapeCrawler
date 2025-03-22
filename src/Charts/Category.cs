using DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal sealed class Category : ICategory
{
    private readonly NumericValue cachedValue;

    internal Category(NumericValue cachedValue)
    {
        this.cachedValue = cachedValue;
    }

    public bool HasMainCategory => false;
    
    public ICategory MainCategory => throw new SCException($"The main category is not available since the chart doesn't have a multi-category. Use {nameof(ICategory.HasMainCategory)} property to check if the main category is available.");

    public string Name
    {
        get => this.cachedValue.InnerText;
        set => this.cachedValue.Text = value;
    }
}