using System.Runtime.CompilerServices;

namespace ShapeCrawler.DevTests.helpers;

public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Init()
    {
        // Initialize the ImageMagick plugin
        VerifyImageMagick.Initialize();

        // Register comparers with a tolerance (threshold).
        // This handles cross-environment font rendering issues.
        VerifyImageMagick.RegisterComparers();
    }
}