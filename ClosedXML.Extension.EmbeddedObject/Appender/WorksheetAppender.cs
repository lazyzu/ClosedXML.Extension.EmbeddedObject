using ClosedXML.Extension.EmbeddedObject.Appender.Sub;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace ClosedXML.Extension.EmbeddedObject.Appender
{
    public class WorksheetAppender
    {
        internal readonly WorksheetPart WorksheetPart;
        internal readonly WorksheetLayoutIdGenerator WorksheetLayoutIdGenerator;
        internal readonly TableParts WorkSheetTableParts;

        private readonly VmlDrawingAppender VmlDrawingAppender;
        private readonly DrawingAppender DrawingAppender;
        private readonly EmbeddedObjectAppender EmbeddedObjectAppender;
        internal readonly WorksheetAppenderUtil Util;

        private EmbeddedObjectShapeIdGenerator EmbeddedObjectShapeIdGenerator;
        private VmlDrawingAppender.VmlDrawingPartInitResult VmlDrawingContext;
        private DrawingAppender.DrawingPartInitResult DrawingContext;
        private EmbeddedObjectAppender.EmbeddedObjectInWorksheetInitResult EmbeddedObjectContext;

        public WorksheetAppender(WorksheetPart worksheetPart, WorksheetLayoutIdGenerator worksheetLayoutIdGenerator)
        {
            this.WorksheetPart = worksheetPart;
            this.WorksheetLayoutIdGenerator = worksheetLayoutIdGenerator;
            this.WorkSheetTableParts = this.WorksheetPart.Worksheet.GetFirstChild<TableParts>();

            this.VmlDrawingAppender = new VmlDrawingAppender(this);
            this.DrawingAppender = new DrawingAppender(this);
            this.EmbeddedObjectAppender = new EmbeddedObjectAppender(this);
            this.Util = new WorksheetAppenderUtil(this);
        }

        public void PrepareEnvContext(int newOleObjectCount)
        {
            AddNamespaceDeclarationToWorksheet(WorksheetPart.Worksheet);
            this.DrawingContext = this.DrawingAppender.InitWorksheetDrawingPart(this.WorksheetPart);
            this.VmlDrawingContext = this.VmlDrawingAppender.InitWorksheetVmlDrawingPart(this.WorksheetPart);
            this.EmbeddedObjectContext = this.EmbeddedObjectAppender.InitEmbeddedObjectInWorksheet(this.WorksheetPart.Worksheet);

            this.EmbeddedObjectShapeIdGenerator = this.WorksheetLayoutIdGenerator.GetShapeIdGenerator(this.WorksheetPart, newOleObjectCount);
        }

        private void AddOleObject(EmbeddedObjectInfo embeddedObjectInfo)
        {
            var addEmbeddedObjectPartContentResult = this.EmbeddedObjectAppender.AddEmbeddedObjectPartContent(WorksheetPart, embeddedObjectInfo);    // Add Embedded Object Content File (OLE Object)

            if (addEmbeddedObjectPartContentResult != null)
            {
                var newShapeId = EmbeddedObjectShapeIdGenerator.Generate();

                var vmlDrawingImagePart = this.VmlDrawingAppender.AddVmlDrawingImagePartToVmlDrawing(VmlDrawingContext.VmlDrawingPart, embeddedObjectInfo, out var vmlDrawingImagePartVmlDrawingId);   // Add icon image file
                this.VmlDrawingAppender.AddVmlDrawingImagePartToWorksheet(WorksheetPart, vmlDrawingImagePart, out var vmlDrawingImagePartWorksheetId);  // Associate icon image file to worksheet
                this.VmlDrawingAppender.AddShapeToVmlDrawingPart(VmlDrawingContext.VmlDrawingPart, embeddedObjectInfo, newShapeId, vmlDrawingImagePartVmlDrawingId);   // Add shape contains icon image for embedded object preview

                this.DrawingAppender.AddDrawingPart(DrawingContext.DrawingPart.WorksheetDrawing, embeddedObjectInfo, newShapeId);  // Add preview drawing

                this.EmbeddedObjectAppender.LinkEmbeddedObjectPartToWorksheet(EmbeddedObjectContext.OleObjects, embeddedObjectInfo, addEmbeddedObjectPartContentResult.WorksheetId, vmlDrawingImagePartWorksheetId, newShapeId); // Link Embedded Object Content File (OLE Object) to worksheet
            }
        }

        private async Task AddOleObjectAsync(EmbeddedObjectInfo embeddedObjectInfo, CancellationToken cancellationToken = default)
        {
            var addEmbeddedObjectPartContentResult = await this.EmbeddedObjectAppender.AddEmbeddedObjectPartContentAsync(WorksheetPart, embeddedObjectInfo, cancellationToken);    // Add Embedded Object Content File (OLE Object)

            if (addEmbeddedObjectPartContentResult != null)
            {
                var newShapeId = EmbeddedObjectShapeIdGenerator.Generate();

                var vmlDrawingImagePart = this.VmlDrawingAppender.AddVmlDrawingImagePartToVmlDrawing(VmlDrawingContext.VmlDrawingPart, embeddedObjectInfo, out var vmlDrawingImagePartVmlDrawingId);   // Add icon image file
                this.VmlDrawingAppender.AddVmlDrawingImagePartToWorksheet(WorksheetPart, vmlDrawingImagePart, out var vmlDrawingImagePartWorksheetId);  // Associate icon image file to worksheet
                this.VmlDrawingAppender.AddShapeToVmlDrawingPart(VmlDrawingContext.VmlDrawingPart, embeddedObjectInfo, newShapeId, vmlDrawingImagePartVmlDrawingId);   // Add shape contains icon image for embedded object preview

                this.DrawingAppender.AddDrawingPart(DrawingContext.DrawingPart.WorksheetDrawing, embeddedObjectInfo, newShapeId);  // Add preview drawing

                this.EmbeddedObjectAppender.LinkEmbeddedObjectPartToWorksheet(EmbeddedObjectContext.OleObjects, embeddedObjectInfo, addEmbeddedObjectPartContentResult.WorksheetId, vmlDrawingImagePartWorksheetId, newShapeId); // Link Embedded Object Content File (OLE Object) to worksheet
            }
        }

        public void AddOleObject(IEnumerable<EmbeddedObjectInfo> embeddedObjectInfos)
        {
            foreach (var embeddedObjectInfo in embeddedObjectInfos) 
            {
                AddOleObject(embeddedObjectInfo);
            }
        }

        public async Task AddOleObjectAsync(IEnumerable<EmbeddedObjectInfo> embeddedObjectInfos, CancellationToken cancellationToken = default)
        {
            foreach (var embeddedObjectInfo in embeddedObjectInfos)
            {
                await AddOleObjectAsync(embeddedObjectInfo, cancellationToken);
            }
        }

        private void AddNamespaceDeclarationToWorksheet(Worksheet worksheet)
        {
            Util.TryToAddNamespaceDeclaration(worksheet, new NamespaceDeclarationInfo[]
            {
                new NamespaceDeclarationInfo("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
                new NamespaceDeclarationInfo("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"),
                new NamespaceDeclarationInfo("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"),
                new NamespaceDeclarationInfo("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
                new NamespaceDeclarationInfo("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"),
            });

            if (worksheet.MCAttributes == null) worksheet.MCAttributes = new MarkupCompatibilityAttributes();
            if (worksheet.MCAttributes.Ignorable?.Value == null) worksheet.MCAttributes.Ignorable = "x14ac";
            else
            {
                var missingIgnorableTarget = worksheet.MCAttributes.Ignorable.Value.IndexOf("x14ac") == -1;
                if (missingIgnorableTarget) worksheet.MCAttributes.Ignorable = string.Join(" ", worksheet.MCAttributes.Ignorable, "x14ac");
            }
        }
    }
}
