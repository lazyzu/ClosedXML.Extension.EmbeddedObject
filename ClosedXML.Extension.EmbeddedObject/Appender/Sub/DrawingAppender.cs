using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace ClosedXML.Extension.EmbeddedObject.Appender.Sub
{
    public class DrawingAppender
    {
        private WorksheetAppender WorksheetAppender;

        internal DrawingAppender(WorksheetAppender WorksheetAppender)
        {
            this.WorksheetAppender = WorksheetAppender;
        }

        internal DrawingPartInitResult InitWorksheetDrawingPart(WorksheetPart worksheetPart)
        {
            var drawingsPart = worksheetPart.GetPartsOfType<DrawingsPart>().FirstOrDefault();
            var needToCreate = drawingsPart == null;
            if (needToCreate) drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            var drawingsPartWorkSheetId = worksheetPart.GetIdOfPart(drawingsPart);

            if (drawingsPart.WorksheetDrawing == null)
            {
                WorksheetDrawing worksheetDrawing = new WorksheetDrawing();
                worksheetDrawing.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                worksheetDrawing.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
                worksheetDrawing.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
                worksheetDrawing.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                drawingsPart.WorksheetDrawing = worksheetDrawing;
            }
            else
            {
                WorksheetDrawing worksheetDrawing = new WorksheetDrawing();
                WorksheetAppender.Util.TryToAddNamespaceDeclaration(worksheetDrawing, new NamespaceDeclarationInfo[]
                {
                    new NamespaceDeclarationInfo("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
                    new NamespaceDeclarationInfo("a14", "http://schemas.microsoft.com/office/drawing/2010/main"),
                    new NamespaceDeclarationInfo("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"),
                    new NamespaceDeclarationInfo("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
                });
            }

            if (needToCreate) LinkDrawingPartToWorksheet(worksheetPart.Worksheet, drawingsPartWorkSheetId);

            return new DrawingPartInitResult
            {
                DrawingPart = drawingsPart,
                WorkSheetId = drawingsPartWorkSheetId,
                Created = needToCreate
            };
        }

        private void LinkDrawingPartToWorksheet(Worksheet worksheet, string drawingPartWorksheetId)
        {
            Drawing drawing = new Drawing() { Id = drawingPartWorksheetId };
            WorksheetAppender.Util.AddAsWorkSheetChild(worksheet, drawing);
        }

        internal class DrawingPartInitResult
        {
            public DrawingsPart DrawingPart { get; init; }
            public string WorkSheetId { get; init; }
            public bool Created { get; init; }
        }

        internal void AddDrawingPart(WorksheetDrawing worksheetDrawing, EmbeddedObjectInfo embeddedObjectInfo, uint vmlDrawingShapeId)
        {
            AlternateContent alternateContent = new AlternateContent();
            alternateContent.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

            AlternateContentChoice alternateContentChoice = new AlternateContentChoice() { Requires = "a14" };
            alternateContentChoice.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            TwoCellAnchor twoCellAnchor = new TwoCellAnchor() { EditAs = EditAsValues.OneCell };

            Xdr.FromMarker fromMarker = new Xdr.FromMarker();
            ColumnId fromColumnId = new ColumnId();
            fromColumnId.Text = embeddedObjectInfo.Dimension.Column.ToString();
            ColumnOffset fromColumnOffset = new ColumnOffset();
            fromColumnOffset.Text = embeddedObjectInfo.Dimension.ColumnFromOffset.ToString();
            RowId fromRowId = new RowId();
            fromRowId.Text = embeddedObjectInfo.Dimension.Row.ToString();
            RowOffset fromRowOffset = new RowOffset();
            fromRowOffset.Text = "0";

            fromMarker.Append(fromColumnId);
            fromMarker.Append(fromColumnOffset);
            fromMarker.Append(fromRowId);
            fromMarker.Append(fromRowOffset);

            Xdr.ToMarker toMarker = new Xdr.ToMarker();
            ColumnId toColumnId = new ColumnId();
            toColumnId.Text = (embeddedObjectInfo.Dimension.Column + embeddedObjectInfo.Dimension.ColumnCountWidth).ToString();
            ColumnOffset toColumnOffset = new ColumnOffset();
            toColumnOffset.Text = embeddedObjectInfo.Dimension.ColumnToOffset.ToString();
            RowId toRowId = new RowId();
            toRowId.Text = (embeddedObjectInfo.Dimension.Row + embeddedObjectInfo.Dimension.RowCountHeight).ToString();
            RowOffset toRowOffset = new RowOffset();
            toRowOffset.Text = "0";

            toMarker.Append(toColumnId);
            toMarker.Append(toColumnOffset);
            toMarker.Append(toRowId);
            toMarker.Append(toRowOffset);

            Shape shape = new Shape() { Macro = "", TextLink = "" };

            NonVisualShapeProperties nonVisualShapeProperties = new NonVisualShapeProperties();

            NonVisualDrawingProperties nonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = (UInt32Value)vmlDrawingShapeId, Name = "", Hidden = true };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension = new A.NonVisualDrawingPropertiesExtension() { Uri = "{63B3BB69-23CF-44E3-9099-C40C66FF867C}" };
            A14.CompatExtension compatExtension1 = new A14.CompatExtension() { ShapeId = $"_x0000_s{vmlDrawingShapeId}" };

            nonVisualDrawingPropertiesExtension.Append(compatExtension1);

            nonVisualDrawingPropertiesExtensionList.Append(nonVisualDrawingPropertiesExtension);

            nonVisualDrawingProperties.Append(nonVisualDrawingPropertiesExtensionList);
            NonVisualShapeDrawingProperties nonVisualShapeDrawingProperties = new NonVisualShapeDrawingProperties();

            nonVisualShapeProperties.Append(nonVisualDrawingProperties);
            nonVisualShapeProperties.Append(nonVisualShapeDrawingProperties);

            ShapeProperties shapeProperties = new ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D = new A.Transform2D();
            A.Offset offset = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents = new A.Extents() { Cx = 0L, Cy = 0L };
            transform2D.Append(offset);
            transform2D.Append(extents);
            shapeProperties.Append(transform2D);

            A.PresetGeometry presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList = new A.AdjustValueList();
            presetGeometry.Append(adjustValueList);
            shapeProperties.Append(presetGeometry);

            A.SolidFill shapeSolidFill = new A.SolidFill();
            A.RgbColorModelHex shapeRgbColorModelHex = new A.RgbColorModelHex() { Val = "FFFFFF", LegacySpreadsheetColorIndex = 65, MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "a14" } };
            shapeSolidFill.Append(shapeRgbColorModelHex);
            shapeProperties.Append(shapeSolidFill);

            A.Outline outline = new A.Outline() { Width = 9525 };
            A.SolidFill outlineSolidFill = new A.SolidFill();
            A.RgbColorModelHex outlineRgbColorModelHex = new A.RgbColorModelHex() { Val = "000000", LegacySpreadsheetColorIndex = 64, MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "a14" } };
            outlineSolidFill.Append(outlineRgbColorModelHex);
            A.Miter outlineMiter = new A.Miter() { Limit = 800000 };
            A.HeadEnd outlineHeadEnd = new A.HeadEnd();
            A.TailEnd outlineTailEnd = new A.TailEnd();
            outline.Append(outlineSolidFill);
            outline.Append(outlineMiter);
            outline.Append(outlineHeadEnd);
            outline.Append(outlineTailEnd);
            shapeProperties.Append(outline);

            shape.Append(nonVisualShapeProperties);
            shape.Append(shapeProperties);
            ClientData clientData = new ClientData();

            twoCellAnchor.Append(fromMarker);
            twoCellAnchor.Append(toMarker);
            twoCellAnchor.Append(shape);
            twoCellAnchor.Append(clientData);

            alternateContentChoice.Append(twoCellAnchor);
            AlternateContentFallback alternateContentFallback = new AlternateContentFallback();

            alternateContent.Append(alternateContentChoice);
            alternateContent.Append(alternateContentFallback);

            worksheetDrawing.Append(alternateContent);
        }
    }
}
