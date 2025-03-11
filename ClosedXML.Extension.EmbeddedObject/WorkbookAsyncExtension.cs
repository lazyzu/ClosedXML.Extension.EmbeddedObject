using ClosedXML.Extension.EmbeddedObject.Appender;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ClosedXML.Extension.EmbeddedObject
{
    public static class WorkbookAsyncExtension
    {
        public static async Task SaveAsWithEmbeddedObjectGenerateAsync(this IXLWorkbook workbook, string file, CancellationToken cancellationToken = default)
        {
            using (var stream = new FileStream(file, FileMode.Create, FileAccess.ReadWrite, FileShare.Read))
            {
                await workbook.SaveAsWithEmbeddedObjectGenerateAsync(stream, cancellationToken);
            }
        }

        public static async Task SaveAsWithEmbeddedObjectGenerateAsync(this IXLWorkbook workbook, string file, bool validate, bool evaluateFormulae = false, CancellationToken cancellationToken = default)
        {
            using (var stream = new FileStream(file, FileMode.Create, FileAccess.ReadWrite, FileShare.Read))
            {
                await workbook.SaveAsWithEmbeddedObjectGenerateAsync(stream, validate, evaluateFormulae, cancellationToken);
            }
        }

        public static async Task SaveAsWithEmbeddedObjectGenerateAsync(this IXLWorkbook workbook, string file, SaveOptions options, CancellationToken cancellationToken = default)
        {
            using (var stream = new FileStream(file, FileMode.Create, FileAccess.ReadWrite, FileShare.Read))
            {
                await workbook.SaveAsWithEmbeddedObjectGenerateAsync(stream, options, cancellationToken);
            }
        }

        public static async Task SaveAsWithEmbeddedObjectGenerateAsync(this IXLWorkbook workbook, Stream stream, CancellationToken cancellationToken = default)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            else
            {
                workbook.SaveAs(stream);
                await EmbeddedObjectGenerateAsync(workbook, stream, cancellationToken);
            }
        }

        public static async Task SaveAsWithEmbeddedObjectGenerateAsync(this IXLWorkbook workbook, Stream stream, bool validate, bool evaluateFormulae = false, CancellationToken cancellationToken = default)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            else
            {
                workbook.SaveAs(stream, validate, evaluateFormulae);
                await EmbeddedObjectGenerateAsync(workbook, stream, cancellationToken);
            }
        }

        public static async Task SaveAsWithEmbeddedObjectGenerateAsync(this IXLWorkbook workbook, Stream stream, SaveOptions options, CancellationToken cancellationToken = default)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            else
            {
                workbook.SaveAs(stream, options);
                await EmbeddedObjectGenerateAsync(workbook, stream, cancellationToken);
            }
        }

        private static async Task EmbeddedObjectGenerateAsync(IXLWorkbook workbook, Stream stream, CancellationToken cancellationToken = default)
        {
            var embeddedObjectsOfWorksheets = workbook.GetWorksheetEmbeddedObjectsMap();
            if (embeddedObjectsOfWorksheets.Count() != 0)
            {
                stream.Position = 0;
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    var worksheetLayoutIdGenerator = new WorksheetLayoutIdGenerator(workbookPart);

                    var sheetPartIdMap = WorkbookExtension.GetSheetPartIdMap(workbookPart, embeddedObjectsOfWorksheets.Keys);

                    foreach (var embeddedObjectsOfWorksheet in embeddedObjectsOfWorksheets)
                    {
                        var sheetname = embeddedObjectsOfWorksheet.Key;
                        var embeddedObjects = embeddedObjectsOfWorksheet.Value.Where(_ => _.OleObject.IsValidCheck()).ToArray();

                        if (embeddedObjects.Length == 0) continue;
                        else
                        {
                            if (sheetPartIdMap.TryGetValue(sheetname, out var sheetId))
                            {
                                var worksheetPart = workbookPart.GetPartById(sheetId) as WorksheetPart;
                                var embeddedObjectAppender = new WorksheetAppender(worksheetPart, worksheetLayoutIdGenerator);
                                embeddedObjectAppender.PrepareEnvContext(embeddedObjects.Length);
                                await embeddedObjectAppender.AddOleObjectAsync(embeddedObjects, cancellationToken);
                            }
                            else { /* Not found sheet in workbook part */ }
                        }
                    }

                    doc.Save();
                }
            }
            else { /*No need to embedded object*/ }
        }
    }
}
