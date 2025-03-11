using ClosedXML.Extension.EmbeddedObject.Appender;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace ClosedXML.Extension.EmbeddedObject
{
    public static class WorksheetExtension
    {
        public static ConditionalWeakTable<IXLWorksheet, List<EmbeddedObjectInfo>> EmbeddedObjectsOfWorksheet
            = new ConditionalWeakTable<IXLWorksheet, List<EmbeddedObjectInfo>>();

        public static List<EmbeddedObjectInfo> GetEmbeddedObjectStore(this IXLWorksheet worksheet)
        {
            if (worksheet == null) return null;
            else return EmbeddedObjectsOfWorksheet.GetOrCreateValue(worksheet);
        }
    }
}
