using ClosedXML.Extension.EmbeddedObject.Appender;
using ClosedXML.Extension.EmbeddedObject.Appender.Icon;
using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.IO;

namespace ClosedXML.Extension.EmbeddedObject.Validation.Appender
{
    public class AppenderTest
    {
        [Test]
        public void AddTxtOleObject()
        {
            var outputTargetPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AddTxtOleObject.xlsx");
            var targetOleContentFilePath = Template.Sample("sample.txt");

            var workbook = new XLWorkbook();
            var helloOleWorksheet = workbook.Worksheets.Add("sampleTxt");

            // 1. Load Object Sote
            var helloOleWorksheetEmbeddedObjectStroe = helloOleWorksheet.GetEmbeddedObjectStore();

            // Render to Cell
            helloOleWorksheetEmbeddedObjectStroe.Add(new EmbeddedObjectInfo
            {
                OleObject = new StreamOleObjectInfo(new FileStream(targetOleContentFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite), targetOleContentFilePath),
                Icon = IconSet.Txt.GetIconInfo(),
                Dimension = DimensionInfo.FromCell(helloOleWorksheet.Cell($"A1"))
            });

            // Render to Cell Center
            helloOleWorksheetEmbeddedObjectStroe.Add(new EmbeddedObjectInfo
            {
                OleObject = new StreamOleObjectInfo(new FileStream(targetOleContentFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite), targetOleContentFilePath),
                Icon = IconSet.Txt.GetIconInfo(),
                Dimension = DimensionInfo.FromCellCenter(helloOleWorksheet.Cell($"A3"))
            });

            // Render to Range
            helloOleWorksheetEmbeddedObjectStroe.Add(new EmbeddedObjectInfo
            {
                OleObject = new StreamOleObjectInfo(new FileStream(targetOleContentFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite), targetOleContentFilePath),
                Icon = IconSet.Txt.GetIconInfo(),
                Dimension = DimensionInfo.FromRange(helloOleWorksheet.Range($"A5", $"A7"))
            });

            // 3. Write the embedded object when saving the workbook
            workbook.SaveAsWithEmbeddedObjectGenerate(outputTargetPath);
        }
    }
}