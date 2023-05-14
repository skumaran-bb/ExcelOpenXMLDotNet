using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelWithOpenXML;

public class ExcelFileService
{
    public ExcelFileService()
    {

    }
    private Stylesheet CreateStylesheet()
    {
        // Creating FONTS
        // *************************************
        Fonts fts = new();

        Font ft = new();
        FontName ftn = new();
        ftn.Val = StringValue.FromString("Calibri");
        ft.FontName = ftn;
        ft.FontSize = new FontSize() { Val = DoubleValue.FromDouble(11) };
        fts.Append(ft); // Adding the font to Fonts[0]

        ft = new Font();
        ftn = new FontName();
        ftn.Val = StringValue.FromString("Calibri");
        ft.FontName = ftn;
        ft.FontSize = new FontSize() { Val = DoubleValue.FromDouble(12) };
        ft.Bold = new Bold() { Val = BooleanValue.FromBoolean(true) };
        fts.Append(ft); // Adding the font to Fonts[1]

        fts.Count = UInt32Value.FromUInt32((uint)fts.ChildElements.Count); // Assigning the fonts count
        // ***** END OF FONTS ********************************


        // Creating FILLS
        // *************************************
        Fills fills = new();

        Fill fill;
        fill = new Fill();
        fill.PatternFill = new PatternFill() { PatternType = PatternValues.None };
        fills.Append(fill); // Adding FILLS[0]

        fill = new Fill();
        fill.PatternFill = new PatternFill() { PatternType = PatternValues.Gray125 };
        fills.Append(fill); // Adding FILLS[1]

        fill = new Fill();
        fill.PatternFill = new PatternFill() { PatternType = PatternValues.Solid };
        fill.PatternFill.ForegroundColor = new ForegroundColor() { Rgb = HexBinaryValue.FromString("00ff9728") };
        fill.PatternFill.BackgroundColor = new BackgroundColor() { Rgb = HexBinaryValue.FromString("00ffc000") };
        fills.Append(fill); // Adding FILLS[2]

        fills.Count = UInt32Value.FromUInt32((uint)fills.ChildElements.Count);
        // ***** END OF FILLS ********************************


        // Creating BORDERS
        // *************************************
        Borders borders = new Borders();

        Border border = new Border();
        border.LeftBorder = new LeftBorder();
        border.RightBorder = new RightBorder();
        border.TopBorder = new TopBorder();
        border.BottomBorder = new BottomBorder();
        border.DiagonalBorder = new DiagonalBorder();
        borders.Append(border); // Adding BORDERS[0]

        border = new Border();
        border.LeftBorder = new LeftBorder();
        border.LeftBorder.Style = BorderStyleValues.Thin;
        border.RightBorder = new RightBorder();
        border.RightBorder.Style = BorderStyleValues.Thin;
        border.TopBorder = new TopBorder();
        border.TopBorder.Style = BorderStyleValues.Thin;
        border.BottomBorder = new BottomBorder();
        border.BottomBorder.Style = BorderStyleValues.Thin;
        border.DiagonalBorder = new DiagonalBorder();
        borders.Append(border); // Adding BORDERS[1]

        borders.Count = UInt32Value.FromUInt32((uint)borders.ChildElements.Count);
        // ***** END OF BORDERS ********************************


        // Creating CELL STYLE FORMAT
        // *************************************
        CellStyleFormats csfs = new CellStyleFormats();

        CellFormat cf = new CellFormat();
        cf.NumberFormatId = 0;
        cf.FontId = 0;
        cf.FillId = 0;
        cf.BorderId = 0;
        csfs.Append(cf);

        csfs.Count = UInt32Value.FromUInt32((uint)csfs.ChildElements.Count);
        // ***** END OF CELL STYLE FORMAT ********************************


        // Creating NUMBERING FORMAT
        // *************************************
        uint iExcelIndex = 164;
        NumberingFormats nfs = new();

        NumberingFormat nfDateTime = new();
        nfDateTime.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
        nfDateTime.FormatCode = StringValue.FromString("dd-MMM-yy");
        //nfDateTime.FormatCode = StringValue.FromString("dd/mm/yyyy hh:mm:ss");
        nfs.Append(nfDateTime); // NUMBERINGFORMATS[0]

        NumberingFormat nf4decimal = new();
        nf4decimal.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
        nf4decimal.FormatCode = StringValue.FromString("#,##0.0000");
        nfs.Append(nf4decimal); // NUMBERINGFORMATS[1]

        // #,##0.00 is also Excel style index 4
        NumberingFormat nf2decimal = new();
        nf2decimal.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
        nf2decimal.FormatCode = StringValue.FromString("#,##0.00");
        nfs.Append(nf2decimal); // NUMBERINGFORMATS[2]

        // @ is also Excel style index 49
        NumberingFormat nfForcedText = new();
        nfForcedText.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++);
        nfForcedText.FormatCode = StringValue.FromString("@");
        nfs.Append(nfForcedText); // NUMBERINGFORMATS[3]

        nfs.Count = UInt32Value.FromUInt32((uint)nfs.ChildElements.Count);

        // ***** END OF NUMBERING FORMAT ********************************


        // Creating CELL FORMATS
        // *************************************
        CellFormats cfs = new CellFormats();

        // Index(0)
        cf = new CellFormat();
        cf.NumberFormatId = 0;
        cf.FontId = 0;
        cf.FillId = 0;
        cf.BorderId = 0;
        cf.FormatId = 0;
        cfs.Append(cf);

        // Index(1)
        cf = new CellFormat();
        cf.NumberFormatId = nfDateTime.NumberFormatId; // DATEIME
        cf.FontId = 0;
        cf.FillId = 0;
        cf.BorderId = 1; // Assigning BORDER[1]
        cf.FormatId = 0;
        cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
        cf.Alignment = new Alignment() { Horizontal = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Center), Vertical = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Center), WrapText = BooleanValue.FromBoolean(false) };
        cfs.Append(cf);

        // Index(2)
        cf = new CellFormat();
        cf.NumberFormatId = nf4decimal.NumberFormatId; // DECIMAL(4)
        cf.FontId = 0;
        cf.FillId = 0;
        cf.BorderId = 0;
        cf.FormatId = 0;
        cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
        cfs.Append(cf);

        // Index(3)
        cf = new CellFormat();
        cf.NumberFormatId = nf2decimal.NumberFormatId; // DECIMAL(2)
        cf.FontId = 0;
        cf.FillId = 0;
        cf.BorderId = 0;
        cf.FormatId = 0;
        cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
        cfs.Append(cf); 

        // Index(4)
        cf = new CellFormat();
        cf.NumberFormatId = nfForcedText.NumberFormatId;
        cf.FontId = 0;
        cf.FillId = 0;
        cf.BorderId = 0;
        cf.FormatId = 0;
        cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
        cfs.Append(cf);

        // Index(5)
        // Header text
        cf = new CellFormat();
        cf.NumberFormatId = nfForcedText.NumberFormatId;
        cf.FontId = 1;
        cf.FillId = 2;
        cf.BorderId = 1;
        cf.FormatId = 0;
        cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
        cf.Alignment = new Alignment() { Horizontal = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Center), Vertical = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Center), WrapText = BooleanValue.FromBoolean(true) };
        cfs.Append(cf);

        // Index(6)
        // Column text
        cf = new CellFormat();
        cf.NumberFormatId = nfForcedText.NumberFormatId;
        cf.FontId = 0;
        cf.FillId = 0;
        cf.BorderId = 1;
        cf.FormatId = 0;
        cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
        cfs.Append(cf);

        // Index(7)
        // Coloured Decimal(2) text
        cf = new CellFormat();
        cf.NumberFormatId = nf2decimal.NumberFormatId;
        cf.FontId = 0;
        cf.FillId = 2;
        cf.BorderId = 0;
        cf.FormatId = 0;
        cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
        cfs.Append(cf);

        // Index(8)
        // Coloured column text
        cf = new CellFormat();
        cf.NumberFormatId = nfForcedText.NumberFormatId;
        cf.FontId = 1; //0
        cf.FillId = 2;
        cf.BorderId = 1;
        cf.FormatId = 0;
        cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
        cfs.Append(cf);

        // Index(9)
        // Column text centered aligned for both horizontal & vertical
        cf = new CellFormat();
        cf.NumberFormatId = nfForcedText.NumberFormatId;
        cf.FontId = 0;
        cf.FillId = 0;
        cf.BorderId = 1;
        cf.FormatId = 0;
        cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
        cf.Alignment = new Alignment() { Horizontal = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Center), Vertical = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Center), WrapText = BooleanValue.FromBoolean(false) };
        cfs.Append(cf);

        // Index(10)
        // column text vertical centered aligned
        cf = new CellFormat();
        cf.NumberFormatId = nfForcedText.NumberFormatId;
        cf.FontId = 0;
        cf.FillId = 0;
        cf.BorderId = 1;
        cf.FormatId = 0;
        cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
        cf.Alignment = new Alignment() { Horizontal = new EnumValue<HorizontalAlignmentValues>(HorizontalAlignmentValues.Left), Vertical = new EnumValue<VerticalAlignmentValues>(VerticalAlignmentValues.Center), WrapText = BooleanValue.FromBoolean(false) };
        cfs.Append(cf);

        cfs.Count = UInt32Value.FromUInt32((uint)cfs.ChildElements.Count);

        // ***** END OF CELL FORMATS ********************************

        // Creating STYLESHEET
        // *************************************
        //Adding Formats, Fonts, Fills, Borders, CellStyleFormats, CellFormats to STYLESHEET
        Stylesheet ss = new();
        ss.Append(nfs);
        ss.Append(fts);
        ss.Append(fills);
        ss.Append(borders);
        ss.Append(csfs);
        ss.Append(cfs);

        CellStyles css = new CellStyles();
        CellStyle cs = new CellStyle();
        cs.Name = StringValue.FromString("Normal");
        cs.FormatId = 0;
        cs.BuiltinId = 0;
        css.Append(cs);
        css.Count = UInt32Value.FromUInt32((uint)css.ChildElements.Count);
        ss.Append(css);

        DifferentialFormats dfs = new DifferentialFormats();
        dfs.Count = 0;
        ss.Append(dfs);

        TableStyles tss = new TableStyles();
        tss.Count = 0;
        tss.DefaultTableStyle = StringValue.FromString("TableStyleMedium9");
        tss.DefaultPivotStyle = StringValue.FromString("PivotStyleLight16");
        ss.Append(tss);

        // *******END of STYLESHEET ******************************

        return ss;
    }

    public void BuildWorkbook(string filename)
    {
        try
        {
            RowData[] rows = new RowData[3];
            rows[0] = new RowData() { EmpCode = 1001, EmpName = "Employee 1", DOB = new DateTime(1982, 08, 19), Salary = 12345.67M, Dept = "Computer" };
            rows[1] = new RowData() { EmpCode = 1002, EmpName = "Employee 2", DOB = new DateTime(1982, 08, 19), Salary = 54321, Dept = "Accounts" };
            rows[2] = new RowData() { EmpCode = 1003, EmpName = "Employee 3", DOB = new DateTime(1982, 08, 19), Salary = 987654.10M, Dept = "Admin" };

            using (SpreadsheetDocument xl = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart wbp = xl.AddWorkbookPart();
                WorksheetPart wsp = wbp.AddNewPart<WorksheetPart>();
                Workbook wb = new Workbook();
                FileVersion fv = new FileVersion();
                fv.ApplicationName = "Microsoft Office Excel";
                Worksheet ws = new Worksheet();
                SheetData sd = new SheetData();

                WorkbookStylesPart wbsp = wbp.AddNewPart<WorkbookStylesPart>();
                wbsp.Stylesheet = CreateStylesheet();
                wbsp.Stylesheet.Save();

                Columns columns = new Columns();
                columns.Append(CreateColumnData(1, 1, 8));
                columns.Append(CreateColumnData(2, 2, 12));
                columns.Append(CreateColumnData(3, 3, 18));
                columns.Append(CreateColumnData(4, 4, 22));
                columns.Append(CreateColumnData(5, 5, 12));
                columns.Append(CreateColumnData(6, 6, 12));

                ws.Append(columns);

                Row r;
                Cell c;

                // header
                r = new Row();
                r.Height = DoubleValue.FromDouble(90);
                r.CustomHeight = BooleanValue.FromBoolean(true);

                c = new Cell();
                c.DataType = CellValues.String;
                c.CellReference = "A1";
                c.CellValue = new CellValue("S. No.");
                c.StyleIndex = 5; // Header style
                r.Append(c);
                

                c = new Cell();
                c.DataType = CellValues.String;
                c.CellReference = "B1";
                c.CellValue = new CellValue("Emp Code");
                c.StyleIndex = 5;
                r.Append(c);

                c = new Cell();
                c.DataType = CellValues.String;
                c.CellReference = "C1";
                c.CellValue = new CellValue("Emp Name");
                c.StyleIndex = 5;
                r.Append(c);

                c = new Cell();
                c.DataType = CellValues.String;
                c.CellReference = "D1";
                c.CellValue = new CellValue("Date of Birth");
                c.StyleIndex = 5;
                r.Append(c);

                c = new Cell();
                c.DataType = CellValues.String;
                c.CellReference = "E1";
                c.CellValue = new CellValue("Salary");
                c.StyleIndex = 5;
                r.Append(c);

                c = new Cell();
                c.DataType = CellValues.String;
                c.CellReference = "F1";
                c.CellValue = new CellValue("Department");
                c.StyleIndex = 5;
                r.Append(c);

                sd.Append(r);


                // content
                int rowNo = 2;

                foreach (var data in rows)
                {
                    r = new Row();
                    r.Height = DoubleValue.FromDouble(30);
                    r.CustomHeight = BooleanValue.FromBoolean(true);

                    c = new Cell();
                    c.DataType = CellValues.String;
                    c.CellReference = $"A{rowNo}";
                    c.CellValue = new CellValue($"{rowNo-1}");
                    c.StyleIndex = 9;
                    r.Append(c);

                    c = new Cell();
                    c.DataType = CellValues.Number;
                    c.CellReference = $"B{rowNo}";
                    c.CellValue = new CellValue(data.EmpCode);
                    c.StyleIndex = 9; // for center alignment
                    r.Append(c);

                    c = new Cell();
                    c.DataType = CellValues.String;
                    c.CellReference = $"C{rowNo}";
                    c.CellValue = new CellValue(data.EmpName);
                    c.StyleIndex = 9; // for center alignment
                    r.Append(c);

                    c = new Cell();
                    c.DataType = CellValues.Date;
                    c.CellReference = $"D{rowNo}";
                    c.CellValue = new CellValue(data.DOB);
                    c.StyleIndex = 1; // for date
                    r.Append(c);

                    c = new Cell();
                    c.DataType = CellValues.Number;
                    c.CellReference = $"E{rowNo}";
                    c.CellValue = new CellValue(data.Salary);
                    c.StyleIndex = 9; // for center alignment
                    r.Append(c);

                    c = new Cell();
                    //c.StyleIndex = 3;
                    c.DataType = CellValues.String;
                    c.CellReference = $"F{rowNo}";
                    c.CellValue = new CellValue(data.Dept);
                    c.StyleIndex = 9; // for center alignment
                    r.Append(c);

                    sd.Append(r);

                    rowNo++;
                }

                ws.Append(sd);
                wsp.Worksheet = ws;
                wsp.Worksheet.Save();
                Sheets sheets = new Sheets();
                Sheet sheet = new Sheet();
                sheet.Name = "Employee Data";
                sheet.SheetId = 1;
                sheet.Id = wbp.GetIdOfPart(wsp);
                sheets.Append(sheet);
                wb.Append(fv);
                wb.Append(sheets);

                xl.WorkbookPart.Workbook = wb;
                xl.WorkbookPart.Workbook.Save();
                xl.Dispose();
            }
        }
        catch (Exception e)
        {
            Console.WriteLine(e.ToString());
            Console.ReadLine();
        }
    }

    private Column CreateColumnData(UInt32 StartColumnIndex, UInt32 EndColumnIndex, double ColumnWidth)
    {
        Column column;
        column = new Column();
        column.Min = StartColumnIndex;
        column.Max = EndColumnIndex;
        column.Width = ColumnWidth;
        column.CustomWidth = true;
        return column;
    }

}
