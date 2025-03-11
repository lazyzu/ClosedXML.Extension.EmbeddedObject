using ClosedXML.Extension.EmbeddedObject.Appender.Sub;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Extension.EmbeddedObject.Appender
{
    public class WorksheetLayoutIdGenerator
    {
        private readonly Queue<int> ValidLayoutIds;
        private Dictionary<WorksheetPart, List<int>> WorksheetLayoutIdMap = new Dictionary<WorksheetPart, List<int>>();

        public WorksheetLayoutIdGenerator(WorkbookPart workbookPart)
        {
            var invalidLayoutIds = new List<int>();

            foreach (var worksheetPart in workbookPart.WorksheetParts)
            {
                var worksheetShapeLayoutIds = VmlDrawingAppender.GetWorksheetShapeLayoutId(worksheetPart).ToList();

                WorksheetLayoutIdMap.Add(worksheetPart, worksheetShapeLayoutIds);
                invalidLayoutIds.AddRange(worksheetShapeLayoutIds);
            }

            var defaultValidIds = FromRange(100, 999, gap: 2);

            if (invalidLayoutIds.Any())
            {
                ValidLayoutIds = new Queue<int>(defaultValidIds.Except(invalidLayoutIds));
            }
            else ValidLayoutIds = new Queue<int>(defaultValidIds);
        }

        public IEnumerable<int> FromRange(int from, int to, int gap)
        {
            for (int i = from; i <= to; i += gap)
            {
                yield return i;
            }
        }

        public int GetNewLayoutId()
        {
            if (ValidLayoutIds.Count == 0) throw new IndexOutOfRangeException("Not able to generate layout id for worksheet, all id were used");
            else return ValidLayoutIds.Dequeue();
        }

        public EmbeddedObjectShapeIdGenerator GetShapeIdGenerator(WorksheetPart worksheetPart, int newOleObjectCount)
        {
            if (WorksheetLayoutIdMap.TryGetValue(worksheetPart, out var layoutIds))
            {
                return new EmbeddedObjectShapeIdGenerator(this, worksheetPart, layoutIds.ToArray(), newOleObjectCount);
            }
            else throw new ArgumentException("Not able to recognize worksheetPart");
        }
    }


    public class EmbeddedObjectShapeIdGenerator
    {
        private readonly WorksheetLayoutIdGenerator WorksheetLayoutIdGenerator;
        private readonly Queue<uint> ValidIds;

        public EmbeddedObjectShapeIdGenerator(WorksheetLayoutIdGenerator worksheetLayoutIdGenerator, WorksheetPart worksheetPart, int[] worksheetLayoutIds, int newOleObjectCount)
        {
            this.WorksheetLayoutIdGenerator = worksheetLayoutIdGenerator;

            var invalidIds = new List<uint>(getExistObjectShapeIdInWorksheet(worksheetPart.Worksheet));
            ValidIds = new Queue<uint>(getValidIds(worksheetLayoutIds, invalidIds));

            var newLayoutIds = new List<int>();
            while (newOleObjectCount > ValidIds.Count())
            {
                var newLayoutId = WorksheetLayoutIdGenerator.GetNewLayoutId();
                foreach (var newValidOleObjectId in getValidIds(new int[] { newLayoutId }, null))
                {
                    ValidIds.Enqueue(newValidOleObjectId);
                }
                newLayoutIds.Add(newLayoutId);
            }
            VmlDrawingAppender.AddWorksheetShapeLayoutId(worksheetPart, newLayoutIds);
        }

        private uint[] getValidIds(int[] worksheetLayoutIds, IEnumerable<uint> invalidIds)
        {
            var fullRange = worksheetLayoutIds.Select(layoutId =>
            {
                var oleObjectBase = (uint)(layoutId * 1000);
                return range(oleObjectBase + 25, oleObjectBase + 999);
            }).SelectMany(id => id)
              .ToArray();

            var invalidIdArray = invalidIds?.ToArray();
            if (invalidIdArray == null) return fullRange;
            if (invalidIdArray.Length == 0) return fullRange;

            return fullRange.Except(invalidIdArray).ToArray();
        }

        private IEnumerable<uint> range(uint from, uint to)
        {
            for (uint i = from; i <= to; i++)
            {
                yield return i;
            }
        }

        public uint Generate()
        {
            if (ValidIds.Count > 0)
            {
                return ValidIds.Dequeue();
            }
            else throw new IndexOutOfRangeException("Not able to generate shape id for ole object, all id were used");
        }

        private IEnumerable<uint> getExistObjectShapeIdInWorksheet(Worksheet worksheet)
        {
            var oleObjects = worksheet.GetFirstChild<OleObjects>();
            if (oleObjects == null) yield break;
            else
            {
                var alternateContents = oleObjects.Elements<AlternateContent>();

                foreach (var alternateContent in alternateContents)
                {
                    var alternateContentChoice = alternateContent?.GetFirstChild<AlternateContentChoice>();
                    var oleObject = alternateContentChoice?.GetFirstChild<OleObject>();
                    if (oleObject != null)
                    {
                        if (uint.TryParse(oleObject.ShapeId, out uint shapeId)) yield return shapeId;
                    }
                }
            }
        }
    }
}
