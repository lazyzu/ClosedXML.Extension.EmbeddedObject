using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace ClosedXML.Extension.EmbeddedObject.Appender.Sub
{
    public class VmlDrawingAppender
    {
        private WorksheetAppender WorksheetAppender;

        internal VmlDrawingAppender(WorksheetAppender worksheetAppender)
        {
            this.WorksheetAppender = worksheetAppender;
        }

        #region Init

        internal VmlDrawingPartInitResult InitWorksheetVmlDrawingPart(WorksheetPart worksheetPart)
        {
            var vmlDrawingPart = worksheetPart.GetPartsOfType<VmlDrawingPart>().FirstOrDefault();
            var needToCreate = vmlDrawingPart == null;
            if (needToCreate) vmlDrawingPart = worksheetPart.AddNewPart<VmlDrawingPart>();
            var vmlDrawingPartWorkSheetId = worksheetPart.GetIdOfPart(vmlDrawingPart);
            InitVmlDrawingXml(vmlDrawingPart);
            if (needToCreate) LinkVmlDrawingPartToWorksheet(worksheetPart.Worksheet, vmlDrawingPartWorkSheetId);

            return new VmlDrawingPartInitResult
            {
                VmlDrawingPart = vmlDrawingPart,
                WorkSheetId = vmlDrawingPartWorkSheetId,
                Created = needToCreate
            };
        }

        private void InitVmlDrawingXml(VmlDrawingPart vmlDrawingPart)
        {
            using (var vmlDrawingXmlStream = vmlDrawingPart.GetStream(FileMode.OpenOrCreate))
            {
                InitVmlDrawingXml(vmlDrawingXmlStream);
            }
        }

        public static void InitVmlDrawingXml(Stream vmlDrawingXmlStream)
        {
            if (vmlDrawingXmlStream.Length == 0)    // Empty File
            {
                using (var writer = new XmlTextWriter(vmlDrawingXmlStream, System.Text.Encoding.UTF8))
                {
                    writer.WriteRaw(@$"<xml xmlns:v=""urn:schemas-microsoft-com:vml""
    xmlns:o=""urn:schemas-microsoft-com:office:office""
    xmlns:x=""urn:schemas-microsoft-com:office:excel"">
    <o:shapelayout v:ext=""edit"">
        <o:idmap v:ext=""edit"" data=""""/>
    </o:shapelayout>
    <v:shapetype id=""_x0000_t75"" coordsize=""21600,21600"" o:spt=""75"" o:preferrelative=""t"" path=""m@4@5l@4@11@9@11@9@5xe"" filled=""f"" stroked=""f"">
        <v:stroke joinstyle=""miter""/>
        <v:formulas>
            <v:f eqn=""if lineDrawn pixelLineWidth 0""/>
            <v:f eqn=""sum @0 1 0""/>
            <v:f eqn=""sum 0 0 @1""/>
            <v:f eqn=""prod @2 1 2""/>
            <v:f eqn=""prod @3 21600 pixelWidth""/>
            <v:f eqn=""prod @3 21600 pixelHeight""/>
            <v:f eqn=""sum @0 0 1""/>
            <v:f eqn=""prod @6 1 2""/>
            <v:f eqn=""prod @7 21600 pixelWidth""/>
            <v:f eqn=""sum @8 21600 0""/>
            <v:f eqn=""prod @7 21600 pixelHeight""/>
            <v:f eqn=""sum @10 21600 0""/>
        </v:formulas>
        <v:path o:extrusionok=""f"" gradientshapeok=""t"" o:connecttype=""rect""/>
        <o:lock v:ext=""edit"" aspectratio=""t""/>
    </v:shapetype>
</xml>");
                    writer.Flush();
                    writer.Close();
                }
            }
            else
            {
                vmlDrawingXmlStream.Position = 0;
                XDocument xDoc = XDocument.Load(vmlDrawingXmlStream);

                // Add Missing Namespace
                var attributes = xDoc.Root.Attributes()
                                           .Where(attr => attr.IsNamespaceDeclaration)
                                           .Select(attr =>
                                           {
                                               return attr.Name.LocalName;
                                           })
                                           .ToArray();

                if (attributes.Contains("v") == false) xDoc.Root.Add(new XAttribute(XNamespace.Xmlns + "v", "urn:schemas-microsoft-com:vml"));
                if (attributes.Contains("o") == false) xDoc.Root.Add(new XAttribute(XNamespace.Xmlns + "o", "urn:schemas-microsoft-com:office:office"));
                if (attributes.Contains("x") == false) xDoc.Root.Add(new XAttribute(XNamespace.Xmlns + "x", "urn:schemas-microsoft-com:office:excel"));

                // Add Missing Shape Layout
                var shapeLayoutXmlName = XName.Get("shapelayout", "urn:schemas-microsoft-com:office:office");
                var hasShapeLayout = xDoc.Root.Elements(shapeLayoutXmlName).Any();

                if (hasShapeLayout == false)
                {
                    var shapeLayout = createShapeLayout();
                    xDoc.Root.Add(XElement.Parse(shapeLayout.OuterXml));
                }

                // Add Missing Shape Type
                var shapeTypeXmlName = XName.Get("shapetype", "urn:schemas-microsoft-com:vml");

                var hasShapeType = xDoc.Root
                                       .Elements(shapeTypeXmlName)
                                       .Any(shapeType =>
                                       {
                                           var idXmlName = XName.Get("id", "");
                                           var shapeTypeId = shapeType.Attributes()
                                                                      .FirstOrDefault(attr => idXmlName.Equals(attr.Name));
                                           return "_x0000_t75".Equals(shapeTypeId?.Value);  // TODO: _x0000_t75 可變動?
                                       });

                if (hasShapeType == false)
                {
                    var sharpType = createShapeType();
                    xDoc.Root.Add(XElement.Parse(sharpType.OuterXml));
                }

                vmlDrawingXmlStream.Position = 0;
                xDoc.Save(vmlDrawingXmlStream);
            }
        }

        private static DocumentFormat.OpenXml.Vml.Office.ShapeLayout createShapeLayout()
        {
            var shapeLayout = new DocumentFormat.OpenXml.Vml.Office.ShapeLayout
            {
                Extension = DocumentFormat.OpenXml.Vml.ExtensionHandlingBehaviorValues.Edit
            };
            shapeLayout.Append(new DocumentFormat.OpenXml.Vml.Office.ShapeIdMap
            {
                Extension = DocumentFormat.OpenXml.Vml.ExtensionHandlingBehaviorValues.Edit,
                Data = ""
            });
            return shapeLayout;
        }

        private static DocumentFormat.OpenXml.Vml.Shapetype createShapeType()
        {
            var shapeType = new DocumentFormat.OpenXml.Vml.Shapetype
            {
                Id = "_x0000_t75",
                CoordinateSize = "21600,21600",
                OptionalNumber = 75,
                PreferRelative = true,
                EdgePath = "m@4@5l@4@11@9@11@9@5xe",
                Filled = false,
                Stroked = false,
            };
            shapeType.Append(new DocumentFormat.OpenXml.Vml.Stroke
            {
                JoinStyle = DocumentFormat.OpenXml.Vml.StrokeJoinStyleValues.Miter
            });
            var formulas = new DocumentFormat.OpenXml.Vml.Formulas();
            formulas.Append(new DocumentFormat.OpenXml.Vml.Formula()
            {
                Equation = "if lineDrawn pixelLineWidth 0"
            });
            formulas.Append(new DocumentFormat.OpenXml.Vml.Formula()
            {
                Equation = "sum @0 1 0"
            });
            formulas.Append(new DocumentFormat.OpenXml.Vml.Formula()
            {
                Equation = "sum 0 0 @1"
            });
            formulas.Append(new DocumentFormat.OpenXml.Vml.Formula()
            {
                Equation = "prod @2 1 2"
            });
            formulas.Append(new DocumentFormat.OpenXml.Vml.Formula()
            {
                Equation = "prod @3 21600 pixelWidth"
            });
            formulas.Append(new DocumentFormat.OpenXml.Vml.Formula()
            {
                Equation = "prod @3 21600 pixelHeight"
            });
            formulas.Append(new DocumentFormat.OpenXml.Vml.Formula()
            {
                Equation = "sum @0 0 1"
            });
            formulas.Append(new DocumentFormat.OpenXml.Vml.Formula()
            {
                Equation = "prod @6 1 2"
            });
            formulas.Append(new DocumentFormat.OpenXml.Vml.Formula()
            {
                Equation = "prod @7 21600 pixelWidth"
            });
            formulas.Append(new DocumentFormat.OpenXml.Vml.Formula()
            {
                Equation = "sum @8 21600 0"
            });
            formulas.Append(new DocumentFormat.OpenXml.Vml.Formula()
            {
                Equation = "prod @7 21600 pixelHeight"
            });
            formulas.Append(new DocumentFormat.OpenXml.Vml.Formula()
            {
                Equation = "sum @10 21600 0"
            });
            shapeType.Append(formulas);
            shapeType.Append(new DocumentFormat.OpenXml.Vml.Path
            {
                AllowExtrusion = false,
                AllowGradientShape = true,
                ConnectionPointType = DocumentFormat.OpenXml.Vml.Office.ConnectValues.Rectangle,
            });
            shapeType.Append(new DocumentFormat.OpenXml.Vml.Office.Lock
            {
                Extension = DocumentFormat.OpenXml.Vml.ExtensionHandlingBehaviorValues.Edit,
                AspectRatio = true
            });

            return shapeType;
        }

        private void LinkVmlDrawingPartToWorksheet(Worksheet worksheet, string vmlDrawingPartWorkSheetId)
        {
            LegacyDrawing legacyDrawing = new LegacyDrawing() { Id = vmlDrawingPartWorkSheetId };
            WorksheetAppender.Util.AddAsWorkSheetChild(worksheet, legacyDrawing);
        }

        internal class VmlDrawingPartInitResult
        {
            public VmlDrawingPart VmlDrawingPart { get; init; }
            public string WorkSheetId { get; init; }

            public bool Created { get; init; }
        }

        #endregion

        internal static IEnumerable<int> GetWorksheetShapeLayoutId(WorksheetPart worksheetPart)
        {
            var vmlDrawingPart = worksheetPart?.GetPartsOfType<VmlDrawingPart>()?.FirstOrDefault();

            if (vmlDrawingPart == null) yield break;
            else
            {
                bool isFileNotFound = false;
                string[] shapeLayoutDataStrValues = new string[0];

                try
                {
                    using (var vmlDrawingXmlStream = vmlDrawingPart.GetStream(FileMode.Open, FileAccess.Read))
                    {
                        var xDoc = XDocument.Load(vmlDrawingXmlStream);

                        var shapeLayoutXmlName = XName.Get("shapelayout", "urn:schemas-microsoft-com:office:office");
                        var shapeLayout = xDoc.Root.Element(shapeLayoutXmlName);

                        var idmapXmlName = XName.Get("idmap", "urn:schemas-microsoft-com:office:office");
                        var idmap = shapeLayout.Element(idmapXmlName);

                        var dataXmlName = XName.Get("data", "");
                        var shapeLayoutDataSettingValue = idmap.Attribute(dataXmlName).Value;

                        shapeLayoutDataStrValues = shapeLayoutDataSettingValue.Split(',');
                    }
                }
                catch (FileNotFoundException)
                {
                    isFileNotFound = true;
                }

                if (isFileNotFound == false)
                {
                    foreach (var shapeLayoutDataStr in shapeLayoutDataStrValues)
                    {
                        if (int.TryParse(shapeLayoutDataStr, out var shapeLayoutData)) yield return shapeLayoutData;
                    }
                }
            }
        }

        internal static void AddWorksheetShapeLayoutId(WorksheetPart worksheetPart, IEnumerable<int> newShapeLayoutIds)
        {
            var _newShapeLayoutIds = newShapeLayoutIds?.ToArray();
            if (_newShapeLayoutIds == null) return;
            if (_newShapeLayoutIds.Length == 0) return;

            var vmlDrawingPart = worksheetPart?.GetPartsOfType<VmlDrawingPart>()?.FirstOrDefault();

            if (vmlDrawingPart != null)
            {
                using (var vmlDrawingXmlStream = vmlDrawingPart.GetStream(FileMode.Open, FileAccess.ReadWrite))
                {
                    var xDoc = XDocument.Load(vmlDrawingXmlStream);

                    var shapeLayoutXmlName = XName.Get("shapelayout", "urn:schemas-microsoft-com:office:office");
                    var shapeLayout = xDoc.Root.Element(shapeLayoutXmlName);

                    var idmapXmlName = XName.Get("idmap", "urn:schemas-microsoft-com:office:office");
                    var idmap = shapeLayout.Element(idmapXmlName);

                    var dataXmlName = XName.Get("data", "");
                    var dataAttribute = idmap.Attribute(dataXmlName);
                    if(string.IsNullOrEmpty(dataAttribute.Value)) dataAttribute.Value = string.Join(",", _newShapeLayoutIds);
                    else dataAttribute.Value = string.Join(",", new string[]
                    {
                        dataAttribute.Value
                    }.Concat(_newShapeLayoutIds.Select(id => id.ToString())));

                    vmlDrawingXmlStream.SetLength(0);
                    xDoc.Save(vmlDrawingXmlStream);
                }
            }
        }

        internal ImagePart AddVmlDrawingImagePartToVmlDrawing(VmlDrawingPart vmlDrawingPart, EmbeddedObjectInfo embeddedObjectInfo, out string vmlDrawingId)
        {
            // Image Part
            var imagePart = vmlDrawingPart.AddNewPart<ImagePart>(embeddedObjectInfo.Icon.ImageContentType, null);

            vmlDrawingId = vmlDrawingPart.GetIdOfPart(imagePart);

            using (var iconImageStream = embeddedObjectInfo.Icon.ImageStreamGetter())
            {
                imagePart.FeedData(iconImageStream);
            }

            return imagePart;
        }

        internal class VmlDrawingImageShapeInfo
        {
            public string Id { get; set; }
            public string ShapeId { get; set; }
        }

        internal void AddVmlDrawingImagePartToWorksheet(WorksheetPart worksheetPart, ImagePart vmlDrawingImagePart, out string workSheetId)
        {
            worksheetPart.AddPart(vmlDrawingImagePart);
            workSheetId = worksheetPart.GetIdOfPart(vmlDrawingImagePart);
        }

        internal void AddShapeToVmlDrawingPart(VmlDrawingPart vmlDrawingPart, EmbeddedObjectInfo embeddedObjectInfo, uint vmlDrawingShapeId, string vmlDrawingImagePartVmlDrawingId)
        {
            using (var vmlDrawingPartStream = vmlDrawingPart.GetStream(FileMode.Open, FileAccess.ReadWrite))
            {
                var document = XDocument.Load(vmlDrawingPartStream);
                AddShapeToVmlDrawingPart(document, embeddedObjectInfo, vmlDrawingShapeId, vmlDrawingImagePartVmlDrawingId);
                vmlDrawingPartStream.SetLength(0);
                document.Save(vmlDrawingPartStream);
            }
        }

        private void AddShapeToVmlDrawingPart(XDocument vmlDrawingDocument, EmbeddedObjectInfo embeddedObjectInfo, uint vmlDrawingShapeId, string vmlDrawingImagePartVmlDrawingId)
        {
            var newShape = new DocumentFormat.OpenXml.Vml.Shape()
            {
                Id = $"_x0000_s{vmlDrawingShapeId}",
                Type = "#_x0000_t75",
                Style = "position:absolute",
                Filled = true,
                FillColor = "window [65]",
                Stroked = true,
                StrokeColor = "windowText [64]",
                InsetMode = DocumentFormat.OpenXml.Vml.Office.InsetMarginValues.Auto
            };
            newShape.Append(new DocumentFormat.OpenXml.Vml.Fill
            {
                Color2 = "window [65]"
            });
            newShape.Append(new DocumentFormat.OpenXml.Vml.ImageData
            {
                Title = "",
                RelId = vmlDrawingImagePartVmlDrawingId
            });

            var clientData = new DocumentFormat.OpenXml.Vml.Spreadsheet.ClientData  // reference: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.spreadsheet.clientdata?view=openxml-2.8.1
            {
                ObjectType = DocumentFormat.OpenXml.Vml.Spreadsheet.ObjectValues.Picture,
            };
            var leftColumn = embeddedObjectInfo.Dimension.Column;   // reference: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.spreadsheet.anchor?view=openxml-2.8.1
            var leftOffset = embeddedObjectInfo.Dimension.ColumnFromOffset;
            var topRow = embeddedObjectInfo.Dimension.Row;
            var topOffset = 0;
            var rightColumn = embeddedObjectInfo.Dimension.Column + embeddedObjectInfo.Dimension.ColumnCountWidth;
            var rightOffset = embeddedObjectInfo.Dimension.ColumnToOffset;
            var bottomRow = embeddedObjectInfo.Dimension.Row + embeddedObjectInfo.Dimension.RowCountHeight;
            var bottomOffset = 0;
            clientData.Append(new DocumentFormat.OpenXml.Vml.Spreadsheet.ResizeWithCells());
            clientData.Append(new DocumentFormat.OpenXml.Vml.Spreadsheet.Anchor($"{leftColumn}, {leftOffset}, {topRow}, {topOffset}, {rightColumn}, {rightOffset}, {bottomRow}, {bottomOffset}"));  // 實際上不塞值似乎也正常運作
            clientData.Append(new DocumentFormat.OpenXml.Vml.Spreadsheet.ClipboardFormat());
            newShape.Append(clientData);

            vmlDrawingDocument.Root.Add(XElement.Parse(newShape.OuterXml));
        }
    }
}
