using ClosedXML.Extension.EmbeddedObject.OleBuilder;
using ClosedXML.Excel;
using PathLib;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ClosedXML.Extension.EmbeddedObject.Appender
{
    public class EmbeddedObjectInfo
    {
        public IOleObjectInfo OleObject { get; init; }

        public IconInfo Icon { get; init; }

        public IDimensionInfo Dimension { get; init; }

        public class DimensionInfo
        {
            public int Column { get; init; }
            public int Row { get; init; }
            public int ColumnCountWidth { get; init; }
            public int RowCountHeight { get; init; }
        }

        public class IconInfo
        {
            public string ImageContentType { get; init; }

            public Func<Stream> ImageStreamGetter { get; init; }
        }
    }

    public interface IDimensionInfo
    {
        int Column { get; }
        long ColumnFromOffset { get; }

        int Row { get; }
        int ColumnCountWidth { get; }
        long ColumnToOffset { get; }
        
        int RowCountHeight { get; }
	}

    public class XLCellDimensionInfo : IDimensionInfo
    {
        public IXLCell TargetCell { get; private init; }

        public XLCellDimensionInfo(IXLCell cell)
        {
            if(cell == null) throw new ArgumentNullException(nameof(cell));

            this.TargetCell = cell;
        }

        public int Column => TargetCell.Address.ColumnNumber - 1;   // 1 Base TO 0 Base
        public int Row => TargetCell.Address.RowNumber - 1; // 1 Base TO 0 Base
        public int ColumnCountWidth => 1;
        public int RowCountHeight => 1;

        public long ColumnFromOffset => 0;

        public long ColumnToOffset => 0;
    }

    public class XLCenterCellDimensionInfo : IDimensionInfo
    {
        public IXLCell TargetCell { get; private init; }

        public XLCenterCellDimensionInfo(IXLCell cell)
        {
            if (cell == null) throw new ArgumentNullException(nameof(cell));

            this.TargetCell = cell;

            var rowHeightInPixel = PixelCalculator.GetRowHightInPixel(this.TargetCell);
            var columnWidthInPixel = PixelCalculator.GetColumnWidthInPixel(this.TargetCell);

            if (columnWidthInPixel > rowHeightInPixel)
            {
                var margin = (columnWidthInPixel - rowHeightInPixel) / 2.0;
                ColumnFromOffset = PixelCalculator.ConvertToEnglishMetricUnits((int)margin, PixelCalculator.DefaultDpi);
                ColumnCountWidth = 0;
                ColumnToOffset = PixelCalculator.ConvertToEnglishMetricUnits((int)(margin + rowHeightInPixel), PixelCalculator.DefaultDpi);
            }
            else
            {
                ColumnFromOffset = 0;
                ColumnCountWidth = 1;
                ColumnToOffset = 0;
            }
        }

        public int Column => TargetCell.Address.ColumnNumber - 1;   // 1 Base TO 0 Base
        public int Row => TargetCell.Address.RowNumber - 1; // 1 Base TO 0 Base
        public int ColumnCountWidth { get; private set; }
        public int RowCountHeight => 1;

        public long ColumnFromOffset { get; private set; }

        public long ColumnToOffset { get; private set; }

        public static class PixelCalculator
        {
            public const int DefaultDpi = 96;
            public const int DefaultRowHeightPoint = 15;

            public static double GetRowHightInPixel(IXLCell cell)
            {
                return GetRowHightInPixel(cell.WorksheetRow().Height);
            }

            public static double GetRowHightInPixel(double rowHeight)
            {
                // RowDpi 數量的 Point = PPI數量的Pixel = 1 Inch
                // Reference: https://techcommunity.microsoft.com/t5/excel/how-to-define-the-size-of-the-cells-in-pixels/m-p/3962742
                return Math.Round(rowHeight * (DefaultDpi / 72.0  /*One point is roughly equivalent to 1/72 of an inch*/));
            }

            public static double GetColumnWidthInPixel(IXLCell cell)
            {
                return GetColumnWidthInPixel(cell.WorksheetColumn().Width);
            }

            public static double GetColumnWidthInPixel(double columnWidth)
            {
                // Reference: https://stackoverflow.com/questions/61041830/convert-excel-column-width-between-characters-unit-and-pixels-points/77053112#77053112
                if (columnWidth == 0) return 0;
                else if (columnWidth < 1) return Math.Round(columnWidth * 12);
                else return Math.Round((columnWidth - 1) * 7 + 12);
            }

            // http://polymathprogrammer.com/2009/10/22/english-metric-units-and-open-xml/
            // http://archive.oreilly.com/pub/post/what_is_an_emu.html
            // https://en.wikipedia.org/wiki/Office_Open_XML_file_formats#DrawingML
            public static Int64 ConvertToEnglishMetricUnits(Int32 pixels, Double resolution)
            {
                return Convert.ToInt64(914400L * pixels / resolution);
            }
        }
    }

    public class XLRangeDimensionInfo : IDimensionInfo
    {
        public IXLRange TargetRange { get; private init; }

        public XLRangeDimensionInfo(IXLRange range)
        {
            if (range == null) throw new ArgumentNullException(nameof(range));

            this.TargetRange = range;

            var firstCell = this.TargetRange.FirstCell();
            this.Column = firstCell.Address.ColumnNumber - 1;   // 1 Base TO 0 Base
            this.Row = firstCell.Address.RowNumber - 1; // 1 Base TO 0 Base

            var lastCell = this.TargetRange.LastCell();
            this.ColumnCountWidth = lastCell.Address.ColumnNumber - firstCell.Address.ColumnNumber + 1;
            this.RowCountHeight = lastCell.Address.RowNumber - firstCell.Address.RowNumber + 1;
        }

        public int Column { get; private init; }
        public int Row { get; private init; }
        public int ColumnCountWidth { get; private init; }
        public int RowCountHeight { get; private init; }

        public long ColumnFromOffset => 0;

        public long ColumnToOffset => 0;
    }

    public static class DimensionInfo
    {
        public static IDimensionInfo FromCellCenter(IXLCell cell) => new XLCenterCellDimensionInfo(cell);
        public static IDimensionInfo FromCell(IXLCell cell) => new XLCellDimensionInfo(cell);
        public static IDimensionInfo FromRange(IXLRange range) => new XLRangeDimensionInfo(range);
    }

    public interface IOleObjectInfo
    {
        string FileName { get; }
        string FilePath { get; }
        string TemporaryPath { get; }

        bool IsContentStreamValid();
        Stream ContentStreamGetter();
    }

    public class StreamOleObjectInfo : IOleObjectInfo
    {
        public string FileName { get; private init; }
        public string FilePath { get; private init; }
        public string TemporaryPath { get; private init; }

        private readonly Stream stream;

        public StreamOleObjectInfo(Stream stream, string filePath, string fileName = null)
        {
            this.stream = stream;
            this.FilePath = filePath;

            this.FileName = fileName;
            if (this.FileName == null)
            {
                var sourceFilePurePath = PathLib.PurePath.Create(filePath);
                this.FileName = sourceFilePurePath.Filename;
            }

            this.TemporaryPath = this.FileName;
        }

        public bool IsContentStreamValid()
        {
            if(this.stream == null || this.stream.Length == 0) return false;
            return true;
        }

        public Stream ContentStreamGetter()
        {
            this.stream.Position = 0;
            return this.stream;
        }
    }

    public static class OleObjectInfoExtension
    {
        public static bool IsValidCheck(this IOleObjectInfo oleObject)
        {
            if (oleObject == null) return false;

            return oleObject.IsContentStreamValid();
        }

        internal static Stream BuildOleObject(this IOleObjectInfo oleObjectInfo)
        {
            if (oleObjectInfo == null) throw new ArgumentNullException(nameof(oleObjectInfo));

            if (oleObjectInfo.IsValidCheck())
            {
                using (var data = oleObjectInfo.ContentStreamGetter())
                {
                    if (data == null) return null;
                    if (data.Length == 0) return null;

                    return OleCompoundFileBuilder.Build(new SimpleCompObjStreamInfo(), new OleNativeStreamInfo
                    {
                        FileName = oleObjectInfo.FileName,
                        FilePath = oleObjectInfo.FilePath,
                        TemporaryPath = oleObjectInfo.TemporaryPath,
                        Data = data
                    });
                }
            }
            else return null;
        }

        internal static async Task<Stream> BuildOleObjectAsync(this IOleObjectInfo oleObjectInfo, CancellationToken cancellationToken = default)
        {
            if (oleObjectInfo == null) throw new ArgumentNullException(nameof(oleObjectInfo));

            if (oleObjectInfo.IsValidCheck())
            {
                using (var data = oleObjectInfo.ContentStreamGetter())
                {
                    if (data == null) return null;
                    if (data.Length == 0) return null;

                    return await OleCompoundFileBuilder.BuildAsync(new SimpleCompObjStreamInfo(), new OleNativeStreamInfo
                    {
                        FileName = oleObjectInfo.FileName,
                        FilePath = oleObjectInfo.FilePath,
                        TemporaryPath = oleObjectInfo.TemporaryPath,
                        Data = data
                    }, cancellationToken);
                }
            }
            else return null;
        }
    }
}
