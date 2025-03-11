using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Extension.EmbeddedObject.Appender.Sub
{
    internal class WorksheetAppenderUtil
    {
        private WorksheetAppender WorksheetAppender;

        public WorksheetAppenderUtil(WorksheetAppender worksheetAppender)
        {
            this.WorksheetAppender = worksheetAppender;
        }

        public Dictionary<string, bool> TryToAddNamespaceDeclaration(OpenXmlElement element, IEnumerable<NamespaceDeclarationInfo> namespaceDeclarationInfos)
        {
            if (namespaceDeclarationInfos == null) return new Dictionary<string, bool>();
            else
            {
                var result = new Dictionary<string, bool>();

                var existPrefixs = element.NamespaceDeclarations
                                          .Select(namespaceDeclare => namespaceDeclare.Key)
                                          .ToArray();

                foreach (var namespaceDeclarationInfo in namespaceDeclarationInfos)
                {
                    var alreadyExist = existPrefixs.Contains(namespaceDeclarationInfo.Prefix);
                    if (alreadyExist == false)
                    {
                        element.AddNamespaceDeclaration(namespaceDeclarationInfo.Prefix, namespaceDeclarationInfo.Uri);
                    }

                    result.Add(namespaceDeclarationInfo.Prefix, !alreadyExist);
                }

                return result;
            }
        }

        public void AddAsWorkSheetChild(Worksheet worksheet, OpenXmlElement element)
        {
            if (WorksheetAppender.WorkSheetTableParts == null) worksheet.Append(element);
            else
            {
                worksheet.InsertBefore(element, WorksheetAppender.WorkSheetTableParts);
            }
        }
    }

    internal class NamespaceDeclarationInfo
    {
        public string Prefix { get; init; }
        public string Uri { get; set; }

        public NamespaceDeclarationInfo(string prefix, string uri)
        {
            Prefix = prefix;
            Uri = uri;
        }
    }
}
