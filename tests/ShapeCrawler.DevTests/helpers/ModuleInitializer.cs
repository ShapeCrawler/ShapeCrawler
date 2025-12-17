using NUnit.Framework;

namespace ShapeCrawler.DevTests;

[SetUpFixture]
public sealed class AssemblySetUp
{
    [OneTimeSetUp]
    public void Init()
    {
        // Initialize the ImageMagick plugin.
        VerifyImageMagick.Initialize();

        // Register comparers with a tolerance (threshold).
        // This handles cross-environment font rendering issues.
        VerifyImageMagick.RegisterComparers();
    }
}