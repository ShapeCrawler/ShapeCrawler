using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideMasters;

internal sealed record SlideLayouts : IReadOnlyList<ISlideLayout>
{
    private readonly SlideMaster parentSlideMaster;
    private readonly P.SlideLayoutIdList slideLayoutIdList;
    private readonly Lazy<List<SlideLayout>> layoutsLazy;

    internal SlideLayouts(SlideMaster parentSlideMaster, P.SlideLayoutIdList slideLayoutIdList)
    {
        this.parentSlideMaster = parentSlideMaster;
        this.slideLayoutIdList = slideLayoutIdList;
        this.layoutsLazy = new Lazy<List<SlideLayout>>(this.ParseSlideLayouts());
    }
    
    private List<SlideLayout> ParseSlideLayouts()
    {
        var rIdList = this.slideLayoutIdList.OfType<P.SlideLayoutId>().Select(layoutId => layoutId.RelationshipId!);
        var layouts = new List<SlideLayout>(rIdList.Count());
        var number = 1;
        foreach (var rId in rIdList)
        {
            SlideLayoutPart sdkLayoutPart = this.parentSlideMaster.SDKLayoutPart(rId.Value!);
            layouts.Add(new SlideLayout(this, sdkLayoutPart, number++));
        }

        return layouts;
    }
    
    internal SlideMaster SlideMaster()
    {
        return this.parentSlideMaster;
    }

    public IEnumerator<ISlideLayout> GetEnumerator()
    {
        throw new System.NotImplementedException();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    public int Count { get; }

    public ISlideLayout this[int index] => this.layoutsLazy.Value[index];
}