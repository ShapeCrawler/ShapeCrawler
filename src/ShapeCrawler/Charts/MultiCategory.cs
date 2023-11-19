using DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal sealed class MultiCategory : ICategory
{
    private readonly NumericValue cachedValue;

    internal MultiCategory(ICategory mainCategory, NumericValue cachedValue)
    {
        this.MainCategory = mainCategory;
        this.cachedValue = cachedValue;
    }

    public bool HasMainCategory => true;
    
    public ICategory MainCategory { get; }

    public string Name
    {
        get => this.cachedValue.InnerText;
        set => this.cachedValue.Text = value;
    }
}