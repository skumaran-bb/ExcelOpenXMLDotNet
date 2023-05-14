using ExcelWithOpenXML;

ExcelFileService excelFileService = new();
excelFileService.BuildWorkbook(@$"C:\Projects\ExcelFile.xlsx");