using ClosedXML.Extension.EmbeddedObject.OleBuilder;
using NUnit.Framework;
using System;
using System.IO;

namespace ClosedXML.Extension.EmbeddedObject.Validation
{
    public class OleBuilderTest
    {
        [Test]
        public void CreateHtmlOleObject()
        {
            var targetFilePath = Template.Sample("sample.txt");
            var outputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "oleObject1.bin");

            using (var embeddedFileContent = new FileStream(targetFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                OleCompoundFileBuilder.Build(new SimpleCompObjStreamInfo(), new OleNativeStreamInfo
                {
                    FileName = Path.GetFileName(targetFilePath),
                    FilePath = targetFilePath,
                    TemporaryPath = Path.Combine(Path.GetTempPath(), Path.GetFileName(targetFilePath)),
                    Data = embeddedFileContent
                }, outputPath);
            }
        }
    }
}
