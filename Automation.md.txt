Folder.GetFiles Folder: $'''C:\\Users\\prath\\Desktop\\Test-20250804T102108Z-1-001\\Test\\Invoices''' FileFilter: $'''*''' IncludeSubfolders: False FailOnAccessDenied: True SortBy1: Folder.SortBy.NoSort SortDescending1: False SortBy2: Folder.SortBy.NoSort SortDescending2: False SortBy3: Folder.SortBy.NoSort SortDescending3: False Files=> Files
Excel.LaunchExcel.LaunchAndOpenUnderExistingProcess Path: $'''C:\\Users\\prath\\Desktop\\Invoice Data.xlsx''' Visible: True ReadOnly: False UseMachineLocale: False Instance=> ExcelInstance
Excel.ClearCellsFromExcel.ClearCells Instance: ExcelInstance StartColumn: 1 StartRow: 2 EndColumn: 16384 EndRow: 1048576
LOOP FOREACH CurrentItem IN Files
    Pdf.ExtractTextFromPDF.ExtractText PDFFile: CurrentItem DetectLayout: False ExtractedText=> ExtractedPDFText
    Text.SplitText.SplitWithDelimiter Text: ExtractedPDFText CustomDelimiter: $'''\\r?\\n''' IsRegEx: True Result=> TextList
    Text.ParseText.RegexParseForFirstOccurrence Text: TextList TextToFind: $'''Invoice Number: (.+)''' StartingPosition: 6 IgnoreCase: False Match=> InvoiceIDText
    Text.ParseText.RegexParseForFirstOccurrence Text: TextList TextToFind: $'''Date of Issue: (.*)''' StartingPosition: 7 IgnoreCase: False Match=> StartDateText
    Text.ParseText.RegexParseForFirstOccurrence Text: TextList TextToFind: $'''Due Date: (.+)''' StartingPosition: 8 IgnoreCase: False Match=> DueDateText
    Text.ParseText.RegexParseForFirstOccurrence Text: TextList TextToFind: $'''NGN\\s[\\d,]+''' StartingPosition: 21 IgnoreCase: False Match=> AmountText
    Text.ParseText.RegexParse Text: TextList TextToFind: $'''Email: (.+)''' StartingPosition: 13 IgnoreCase: False OccurrencePositions=> EmailPositions Matches=> EmailText
    Text.SplitText.SplitWithDelimiter Text: StartDateText CustomDelimiter: $'''Date of Issue:''' IsRegEx: False Result=> StartDateValue
    Text.SplitText.SplitWithDelimiter Text: DueDateText CustomDelimiter: $'''Due Date:''' IsRegEx: False Result=> DueDateValue
    Text.SplitText.SplitWithDelimiter Text: InvoiceIDText CustomDelimiter: $'''Invoice Number:''' IsRegEx: False Result=> InvoiceIDValue
    Text.SplitText.SplitWithDelimiter Text: EmailText CustomDelimiter: $'''Email:''' IsRegEx: False Result=> EmailValue
    Text.SplitText.SplitWithDelimiter Text: AmountText CustomDelimiter: $'''TOTAL DUE: NGN''' IsRegEx: False Result=> AmountValue
    SET Client TO TextList[10]
    SET InvoiceID TO InvoiceIDValue[1]
    SET StartDate TO StartDateValue[1]
    SET DueDate TO InvoiceIDValue[1]
    SET Amount TO AmountValue[1]
    SET Email TO EmailValue[2]
    Excel.GetFirstFreeRowOnColumn Instance: ExcelInstance Column: $'''A''' FirstFreeRowOnColumn=> FirstFreeRowOnColumn
    Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: InvoiceID Column: $'''A''' Row: FirstFreeRowOnColumn
    Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: Client Column: $'''B''' Row: FirstFreeRowOnColumn
    Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: Email Column: $'''C''' Row: FirstFreeRowOnColumn
    Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: StartDate Column: $'''D''' Row: FirstFreeRowOnColumn
    Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: DueDate Column: $'''E''' Row: FirstFreeRowOnColumn
    Excel.WriteToExcel.WriteCell Instance: ExcelInstance Value: Amount Column: $'''F''' Row: FirstFreeRowOnColumn
END
Excel.CloseExcel.CloseAndSave Instance: ExcelInstance
