using System.IO;

namespace ShapeCrawler;

internal interface ICopyablePresentation : IPresentationProperties
{
    void Copy(string path);

    void Copy(Stream stream);
}