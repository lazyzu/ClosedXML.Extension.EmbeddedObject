using ClosedXML.Extension.EmbeddedObject.Appender;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ClosedXML.Extension.EmbeddedObject
{
    public static class WorkbookExtension
    {
        public static void SaveAsWithEmbeddedObjectGenerate(this IXLWorkbook workbook, string file)
        {
            using (var stream = new FileStream(file, FileMode.Create, FileAccess.ReadWrite, FileShare.Read))
            {
                workbook.SaveAsWithEmbeddedObjectGenerate(stream);
            }
        }

        public static void SaveAsWithEmbeddedObjectGenerate(this IXLWorkbook workbook, string file, bool validate, bool evaluateFormulae = false)
        {
            using (var stream = new FileStream(file, FileMode.Create, FileAccess.ReadWrite, FileShare.Read))
            {
                workbook.SaveAsWithEmbeddedObjectGenerate(stream, validate, evaluateFormulae);
            }
        }

        public static void SaveAsWithEmbeddedObjectGenerate(this IXLWorkbook workbook, string file, SaveOptions options)
        {
            using (var stream = new FileStream(file, FileMode.Create, FileAccess.ReadWrite, FileShare.Read))
            {
                workbook.SaveAsWithEmbeddedObjectGenerate(stream, options);
            }
        }

        public static void SaveAsWithEmbeddedObjectGenerate(this IXLWorkbook workbook, Stream stream)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            else
            {
                workbook.SaveAs(stream);
                EmbeddedObjectGenerate(workbook, stream);
            }
        }

        public static void SaveAsWithEmbeddedObjectGenerate(this IXLWorkbook workbook, Stream stream, bool validate, bool evaluateFormulae = false)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            else
            {
                workbook.SaveAs(stream, validate, evaluateFormulae);
                EmbeddedObjectGenerate(workbook, stream);
            }
        }

        public static void SaveAsWithEmbeddedObjectGenerate(this IXLWorkbook workbook, Stream stream, SaveOptions options)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            else
            {
                workbook.SaveAs(stream, options);
                EmbeddedObjectGenerate(workbook, stream);
            }
        }

        private static void EmbeddedObjectGenerate(IXLWorkbook workbook, Stream stream)
        {
            var embeddedObjectsOfWorksheets = workbook.GetWorksheetEmbeddedObjectsMap();
            if (embeddedObjectsOfWorksheets.Count() != 0)
            {
                stream.Position = 0;
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, true))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    var worksheetLayoutIdGenerator = new WorksheetLayoutIdGenerator(workbookPart);

                    var sheetPartIdMap = GetSheetPartIdMap(workbookPart, embeddedObjectsOfWorksheets.Keys);

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
                                embeddedObjectAppender.AddOleObject(embeddedObjects);
                            }
                            else { /* Not found sheet in workbook part */ }
                        }
                    }

                    doc.Save();
                }
            }
            else { /*No need to embedded object*/ }
        }

        internal static Dictionary<string, string> GetSheetPartIdMap(WorkbookPart workbookPart, IEnumerable<string> sheetNames)
        {
            var _sheetNames = sheetNames?.ToArray();

            if (_sheetNames == null) return new Dictionary<string, string>();
            if (_sheetNames.Length == 0) return new Dictionary<string, string>();

            return workbookPart.Workbook
                               .Sheets
                               .Cast<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                               .Where(sheet => _sheetNames.Contains(sheet.Name.Value))
                               .ToDictionary(sheet => sheet.Name.Value
                                           , sheet => sheet.Id.Value);
        }

        public static Dictionary<string, EmbeddedObjectInfo[]> GetWorksheetEmbeddedObjectsMap(this IXLWorkbook workbook)
        {
            if (workbook == null) return new Dictionary<string, EmbeddedObjectInfo[]>();
            else
            {
                var result = new Dictionary<string, EmbeddedObjectInfo[]>();
                foreach (var worksheet in workbook.Worksheets)
                {
                    var worksheetEmbeddedObjects = worksheet.GetEmbeddedObjectStore();
                    result.Add(worksheet.Name, worksheetEmbeddedObjects.ToArray());
                }
                return result;
            }
        }
    }
}
