About ClosedXML.Extension.EmbeddedObject:
=================
`ClosedXML.Extension.EmbeddedObject` is a .NET library for embedding OLE objects when using [ClosedXML](https://github.com/ClosedXML/ClosedXML)
And currently, it only supports embedding and does not support reading, updating, or deleting actions.

> [!NOTE]
> Supports ClosedXML versions 0.96.0 and later, except for 0.103.x and 0.104.x due to the [not strongly named](https://github.com/ClosedXML/ClosedXML/issues/2423) issue.

Using ClosedXML.Extension.EmbeddedObject:
=================
* Basic usage
```
public void AddTxtOleObject()
{
    var outputTargetPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AddTxtOleObject.xlsx");
    var targetOleContentFilePath = Template.Sample("sample.txt");

    var workbook = new XLWorkbook();
    var helloOleWorksheet = workbook.Worksheets.Add("sampleTxt");

    // 1. Load Object Sote
    var helloOleWorksheetEmbeddedObjectStroe = helloOleWorksheet.GetEmbeddedObjectStore();

    // 2. Record the information of the object to be embedded in Store
    helloOleWorksheetEmbeddedObjectStroe.Add(new EmbeddedObjectInfo
    {
        OleObject = new StreamOleObjectInfo(new FileStream(targetOleContentFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite), targetOleContentFilePath),
        Icon = IconSet.Txt.GetIconInfo(),
        Dimension = DimensionInfo.FromCell(helloOleWorksheet.Cell($"A1"))
    });

    // 3. Write the embedded object when saving the workbook
    workbook.SaveAsWithEmbeddedObjectGenerate(outputTargetPath);
}
```

* `OleObject` can implement `IOleObjectInfo` for customization, such as the implementation of `StreamOleObjectInfo` as follows:
```
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
```

* The rendering position can be customized by implementing `IDimensionInfo`.
Currently, three rendering methods are supported: FromCell, FromCellCenter, and FromRange.
The implementation of `XLRangeDimensionInfo` is as follows:
```
// Default supported rendering methods
DimensionInfo.FromCell(helloOleWorksheet.Cell($"A1"))
DimensionInfo.FromCellCenter(helloOleWorksheet.Cell($"A3"))
DimensionInfo.FromRange( helloOleWorksheet.Range($"A5", $"A7"))

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
```

* Icon can initialize an `IconInfo` object to customize settings, as shown in the following example.
```
new EmbeddedObjectInfo.IconInfo
{
    ImageContentType = "image/x-emf",
    ImageStreamGetter = getEmbeddedResourceIcon("text.emf")
};

private static Stream getEmbeddedResourceIcon(string fileName)
{
    return _assembly.Value.GetManifestResourceStream($"{_namespace.Value}.{fileName}");
}
```
