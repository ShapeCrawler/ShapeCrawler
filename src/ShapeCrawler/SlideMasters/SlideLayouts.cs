using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideMasters;

internal sealed record SlideLayouts : IReadOnlyList<ISlideLayout>
{
    private readonly SlideMasterPart sdkSlideMasterPart;

    internal SlideLayouts(SlideMasterPart sdkSlideMasterPart)
    {
        this.sdkSlideMasterPart = sdkSlideMasterPart;
    }
    
    private List<ISlideLayout> LayoutList()
    {
        var rIdList = this.sdkSlideMasterPart.SlideMaster.SlideLayoutIdList!.OfType<P.SlideLayoutId>().Select(layoutId => layoutId.RelationshipId!);
        var layouts = new List<ISlideLayout>(rIdList.Count());
        var number = 1;
        foreach (var rId in rIdList)
        {
            var sdkLayoutPart = (SlideLayoutPart)this.sdkSlideMasterPart.GetPartById(rId.Value!);
            var slideMaster = new SlideMaster(this.sdkSlideMasterPart);
            layouts.Add(new SlideLayout(sdkLayoutPart));
        }
    
        return layouts;
    }

    public int Count => this.LayoutList().Count;

    public ISlideLayout this[int index] => this.LayoutList()[index];
    
    public IEnumerator<ISlideLayout> GetEnumerator()
    {
        return this.LayoutList().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }
}