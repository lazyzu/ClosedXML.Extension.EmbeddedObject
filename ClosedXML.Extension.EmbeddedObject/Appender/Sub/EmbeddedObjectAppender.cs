using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;
using System.Threading.Tasks;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace ClosedXML.Extension.EmbeddedObject.Appender.Sub
{
    public class EmbeddedObjectAppender
    {
        private WorksheetAppender WorksheetAppender;

        internal EmbeddedObjectAppender(WorksheetAppender embeddedObjectAppender)
        {
            this.WorksheetAppender = embeddedObjectAppender;
        }

        internal EmbeddedObjectInWorksheetInitResult InitEmbeddedObjectInWorksheet(Worksheet worksheet)
        {
            var oleObjects = worksheet.GetFirstChild<OleObjects>();
            if (oleObjects == null)
            {
                oleObjects = new OleObjects();
                WorksheetAppender.Util.AddAsWorkSheetChild(worksheet, oleObjects);
                return new EmbeddedObjectInWorksheetInitResult
                {
                    OleObjects = oleObjects,
                    Created = true
                };
            }
            else return new EmbeddedObjectInWorksheetInitResult
            {
                OleObjects = oleObjects,
                Created = false
            };
        }

        internal class EmbeddedObjectInWorksheetInitResult
        {
            public OleObjects OleObjects { get; init; }
            public bool Created { get; init; }
        }

        internal AddEmbeddedObjectPartContentResult AddEmbeddedObjectPartContent(WorksheetPart worksheet, EmbeddedObjectInfo embeddedObjectInfo)
        {
            using (var oleObjectStream = embeddedObjectInfo.OleObject.BuildOleObject())
            {
                if (oleObjectStream == null) return null;
                else
                {
                    var embeddedObjectPart = worksheet.AddNewPart<EmbeddedObjectPart>("application/vnd.openxmlformats-officedocument.oleObject", null);
                    embeddedObjectPart.FeedData(oleObjectStream);

                    return new AddEmbeddedObjectPartContentResult
                    {
                        EmbeddedObjectPart = embeddedObjectPart,
                        WorksheetId = worksheet.GetIdOfPart(embeddedObjectPart)
                    };
                }
            }
        }
        
        internal async Task<AddEmbeddedObjectPartContentResult> AddEmbeddedObjectPartContentAsync(WorksheetPart worksheet, EmbeddedObjectInfo embeddedObjectInfo, CancellationToken cancellationToken = default)
        {
            using (var oleObjectStream = await embeddedObjectInfo.OleObject.BuildOleObjectAsync(cancellationToken))
            {
                if (oleObjectStream == null) return null;
                else
                {
                    var embeddedObjectPart = worksheet.AddNewPart<EmbeddedObjectPart>("application/vnd.openxmlformats-officedocument.oleObject", null);
                    embeddedObjectPart.FeedData(oleObjectStream);
                    return new AddEmbeddedObjectPartContentResult
                    {
                        EmbeddedObjectPart = embeddedObjectPart,
                        WorksheetId = worksheet.GetIdOfPart(embeddedObjectPart)
                    };
                }
            }
        }

        internal class AddEmbeddedObjectPartContentResult
        {
            public EmbeddedObjectPart EmbeddedObjectPart { get; init; }
            public string WorksheetId { get; init; }
        }

        internal void LinkEmbeddedObjectPartToWorksheet(OleObjects oleObjects, EmbeddedObjectInfo embeddedObjectInfo, string embeddedObjectPartWorksheetId, string vmlDrawingImagePartWorksheetId, uint vmlDrawingShapeId)
        {
            AlternateContent alternateContent = new AlternateContent();
            alternateContent.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            AlternateContentChoice alternateContentChoice = new AlternateContentChoice() { Requires = "x14" };

            OleObject oleObjectAlternateContentChoice = new OleObject()
            {
                ProgId = "封裝程式殼層物件",
                DataOrViewAspect = DataViewAspectValues.DataViewAspectIcon,
                ShapeId = (UInt32Value)vmlDrawingShapeId,
                Id = embeddedObjectPartWorksheetId
            };

            EmbeddedObjectProperties oleObjectAlternateContentChoiceProperties = new EmbeddedObjectProperties()
            {
                DefaultSize = false,
                Print = false,
                AutoPict = false,
                AltText = "",
                Id = vmlDrawingImagePartWorksheetId
            };

            ObjectAnchor objectAnchor = new ObjectAnchor() { MoveWithCells = true };

            FromMarker fromMarker = new FromMarker();
            Xdr.ColumnId fromColumnId = new Xdr.ColumnId();
            fromColumnId.Text = embeddedObjectInfo.Dimension.Column.ToString();
            Xdr.ColumnOffset fromColumnOffset = new Xdr.ColumnOffset();
            fromColumnOffset.Text = embeddedObjectInfo.Dimension.ColumnFromOffset.ToString();
            Xdr.RowId fromRowId = new Xdr.RowId();
            fromRowId.Text = embeddedObjectInfo.Dimension.Row.ToString();
            Xdr.RowOffset fromRowOffset = new Xdr.RowOffset();
            fromRowOffset.Text = "0";

            fromMarker.Append(fromColumnId);
            fromMarker.Append(fromColumnOffset);
            fromMarker.Append(fromRowId);
            fromMarker.Append(fromRowOffset);

            ToMarker toMarker = new ToMarker();
            Xdr.ColumnId toColumnId = new Xdr.ColumnId();
            toColumnId.Text = (embeddedObjectInfo.Dimension.Column + embeddedObjectInfo.Dimension.ColumnCountWidth).ToString();
            Xdr.ColumnOffset toColumnOffset = new Xdr.ColumnOffset();
            toColumnOffset.Text = embeddedObjectInfo.Dimension.ColumnToOffset.ToString();
            Xdr.RowId toRowId = new Xdr.RowId();
            toRowId.Text = (embeddedObjectInfo.Dimension.Row + embeddedObjectInfo.Dimension.RowCountHeight).ToString();
            Xdr.RowOffset toRowOffset = new Xdr.RowOffset();
            toRowOffset.Text = "0";

            toMarker.Append(toColumnId);
            toMarker.Append(toColumnOffset);
            toMarker.Append(toRowId);
            toMarker.Append(toRowOffset);

            objectAnchor.Append(fromMarker);
            objectAnchor.Append(toMarker);

            oleObjectAlternateContentChoiceProperties.Append(objectAnchor);
            oleObjectAlternateContentChoice.Append(oleObjectAlternateContentChoiceProperties);
            alternateContentChoice.Append(oleObjectAlternateContentChoice);

            AlternateContentFallback alternateContentFallback = new AlternateContentFallback();
            OleObject oleObjectAlternateContentFallback = new OleObject()
            {
                ProgId = "封裝程式殼層物件",
                DataOrViewAspect = DataViewAspectValues.DataViewAspectIcon,
                ShapeId = (UInt32Value)vmlDrawingShapeId,
                Id = embeddedObjectPartWorksheetId
            };
            alternateContentFallback.Append(oleObjectAlternateContentFallback);

            alternateContent.Append(alternateContentChoice);
            alternateContent.Append(alternateContentFallback);

            oleObjects.Append(alternateContent);
        }
    }
}
