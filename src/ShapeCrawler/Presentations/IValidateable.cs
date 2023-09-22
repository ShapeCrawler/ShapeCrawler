using System.IO;

namespace ShapeCrawler;

internal interface IValidateable : IPresentationProperties
{
    void Validate();
    void CopyTo(string path);
    void CopyTo(Stream stream);
}