using System.IO;

namespace ShapeCrawler;

internal interface IValidateable : IPresentationProperties
{
    void Validate();
    void Copy(string path);
    void Copy(Stream stream);
}