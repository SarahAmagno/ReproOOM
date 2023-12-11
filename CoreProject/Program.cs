using GemBox.Spreadsheet;

SpreadsheetInfo.SetLicense("INSERT LICENSE KEY HERE");

const string outputDirectory = @"..\..\..\OutputFiles";
const string inputPath = "InputFile.xlsx";

for (var i = 0; i < 15; i++)
{
    var spreadSheet = ExcelFile.Load(inputPath, new XlsxLoadOptions
    {
        StreamingMode = LoadStreamingMode.Read
    });
    
    var pages = spreadSheet.GetPaginator(new PaginatorOptions { SelectionType = SelectionType.EntireFile }).Pages;
        
    for (var index = 0; index < pages.Count; index++)
    {
        var fileName = $"Spreadsheet-{i}_Page-{index}.bmp";
            
        var page = pages[index];
            
        using var stream = new FileStream(Path.Combine(outputDirectory, fileName), FileMode.Create);
            
        page.Save(stream, new ImageSaveOptions(ImageSaveFormat.Bmp));
    }
}