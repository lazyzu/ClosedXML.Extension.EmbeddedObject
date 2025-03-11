using NUnit.Framework;
using System.IO;

namespace ClosedXML.Extension.EmbeddedObject.Validation
{
    public static class Template
    {
        public static string Folder => Path.Combine(TestContext.CurrentContext.TestDirectory, @"..\..\..\Template\");

        public static string Sample(string referfencePath) => Path.Combine(Folder, referfencePath);

        public static byte[] SampleByteArrayContent(string referfencePath)
        {
            return File.ReadAllBytes(Sample(referfencePath));
        }

        public static Stream SampleStreamContent(string referfencePath)
        {
            return new FileStream(referfencePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }
    }
}
