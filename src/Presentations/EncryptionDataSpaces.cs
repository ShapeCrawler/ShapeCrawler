using System.IO;
using System.Text;
using OpenMcdf;

namespace ShapeCrawler.Presentations;

// MS-OFFCRYPTO 2.2.6: map the encrypted package to the Office encryption transform.
internal static class EncryptionDataSpaces
{
    internal static void Write(RootStorage root)
    {
        var spaces = root.CreateStorage("\u0006DataSpaces");
        using (var writer = Writer(spaces, "Version"))
        {
            WriteString(writer, "Microsoft.Container.DataSpaces");
            WriteVersions(writer);
        }

        using (var writer = Writer(spaces, "DataSpaceMap"))
        {
            writer.Write(8);
            writer.Write(1);
            writer.Write(104);
            writer.Write(1);
            writer.Write(0);
            WriteString(writer, "EncryptedPackage");
            WriteString(writer, "StrongEncryptionDataSpace");
        }

        var info = spaces.CreateStorage("DataSpaceInfo");
        using (var writer = Writer(info, "StrongEncryptionDataSpace"))
        {
            writer.Write(8);
            writer.Write(1);
            WriteString(writer, "StrongEncryptionTransform");
        }

        var transform = spaces.CreateStorage("TransformInfo").CreateStorage("StrongEncryptionTransform");
        using (var writer = Writer(transform, "\u0006Primary"))
        {
            writer.Write(88);
            writer.Write(1);
            WriteString(writer, "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}");
            WriteString(writer, "Microsoft.Container.EncryptionTransform");
            WriteVersions(writer);
            writer.Write(0);
            writer.Write(0);
            writer.Write(0);
            writer.Write(4);
        }
    }

    private static BinaryWriter Writer(Storage storage, string name) => new(storage.CreateStream(name), Encoding.Unicode);

    private static void WriteVersions(BinaryWriter writer)
    {
        writer.Write(1);
        writer.Write(1);
        writer.Write(1);
    }

    private static void WriteString(BinaryWriter writer, string value)
    {
        var bytes = Encoding.Unicode.GetBytes(value);
        writer.Write(bytes.Length);
        writer.Write(bytes);
        writer.Write(new byte[(4 - bytes.Length % 4) % 4]);
    }
}
