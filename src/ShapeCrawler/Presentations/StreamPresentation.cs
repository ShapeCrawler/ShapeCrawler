using System.IO;

namespace ShapeCrawler;

internal sealed class StreamPresentation : IValidateable
{
    private readonly PresentationCore presentationCore;
    private readonly Stream userStream;

    internal StreamPresentation(Stream userStream)
    {
        this.userStream = userStream;
        var internalStream = new MemoryStream();
        this.userStream.Position = 0;
        userStream.CopyTo(internalStream);
        this.presentationCore = new PresentationCore(internalStream);
    }

    public ISlides Slides => this.presentationCore.Slides;
    
    public decimal SlideWidth
    {
        get => this.presentationCore.SlideWidth;
        set => this.presentationCore.SlideWidth = value;
    }
    
    public decimal SlideHeight
    {
        get => this.presentationCore.SlideHeight;
        set => this.presentationCore.SlideHeight = value;
    }
    
    public ISlideMasterCollection SlideMasters => this.presentationCore.SlideMasters;
    
    public ISections Sections => this.presentationCore.Sections;
    
    public IFooter Footer => this.presentationCore.Footer;
    
    public byte[] AsByteArray() => this.presentationCore.AsByteArray();
    
    
    public void Save() => this.presentationCore.CopyTo(this.userStream);
    
    void IValidateable.Validate() => this.presentationCore.Validate();
    
    public void CopyTo(string path) => this.presentationCore.CopyTo(path);
    
    public void CopyTo(Stream stream) => this.presentationCore.CopyTo(stream);
}