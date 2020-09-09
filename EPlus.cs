using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.FormulaParsing;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Reflection;
using System.Data;

namespace ExportExcelHuge
{
    public class EPlusExcel
    {
        /// <summary>
        /// Templete : If header data required as  First Row data
        ///                                        Second Row data
        ///                                        
        ///                    send the template string as "First Row data  € Second Row data"  -- seperate each line text with €
        /// </summary>
        /// <param name="template"></param>
        /// <param name="dtObject"></param>
        /// <param name="sheetName"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public string BuildExcelSheetWithTemplete(string template, DataTable dtObject, string sheetName, string fileName)
        {

            // this method supports for exporting only less than or equal to 26 columns
            using (ExcelPackage pck = new ExcelPackage())
            {
                var DT = dtObject;

                //Create the worksheet;
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add(sheetName);

                // "€" seperated templete.
                int staticDataLineCnt = template.Split('€').Length;   // number of static data lines
                int totalColumns = DT.Columns.Count; // total number of columns in export list
                string lastColChar = Convert.ToChar(65 + (totalColumns - 1)).ToString();  // last column name
                for (int i = 1; i <= template.Split('€').Length; i++)
                {
                    string cellValue = "A" + i + ":" + lastColChar + i;
                    ws.Cells[cellValue].Merge = true;
                    ws.Cells[cellValue].Style.Font.Bold = true;
                    ws.Cells[cellValue].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    ws.Cells[cellValue].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    ws.Cells[cellValue].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[cellValue].LoadFromText(template.Split('€')[i - 1]);
                }

                var dataRange = ws.Cells["A" + (staticDataLineCnt + 1)].LoadFromDataTable(DT, true);


                var ss = ws.Cells;
                var pp = ss.Take(DT.Columns.Count);

                string cellValueNew = "A" + (staticDataLineCnt + 1) + ":" + lastColChar + (staticDataLineCnt + 1);
                using (ExcelRange rng = ws.Cells[cellValueNew])
                {
                    rng.Style.Font.Bold = true;
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;                      //Set Pattern for the background to Solid
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));  //Set color to dark blue
                    rng.Style.Font.Color.SetColor(Color.White);
                    rng.AutoFilter = true;
                    rng.AutoFitColumns(45, 150);
                }

                #region PERSISTS FORMAT DATTYPE
                if (DT.Rows.Count > 0)
                {
                    for (int i = 0; i < DT.Columns.Count; i++)
                    {
                        switch (DT.Columns[i].DataType.Name)
                        {

                            case "DateTime": ws.Cells[2, i + 1, dataRange.End.Row, i + 1].Style.Numberformat.Format = "MM-dd-yyyy hh:mm"; break;
                            case "TimeSpan": //ws.Cells[2, i + 1, dataRange.End.Row, i + 1].Style.Numberformat.Format = "hh:mm"; break;
                            case "Decimal":
                            case "Double": //ws.Cells[2, i + 1, dataRange.End.Row, i + 1].Style.Numberformat.Format = "0:0.##"; break;
                            case "SByte":
                            case "Single":
                            case "String":
                            case "Boolean":
                            case "Byte":
                            case "Char":
                            case "Int16":
                            case "Int32":
                            case "Int64":
                            case "UInt16":
                            case "UInt32":
                            case "UInt64": break;//ws.Cells[2, i + 1, dataRange.End.Row, i + 1].Style.Numberformat.Format = "@"; break;
                            default: break;
                        }

                    }
                }
                #endregion

                pck.SaveAs(new FileInfo(AppKeyHelper.GetKey("EXPORT_FILE_PATH") + fileName + ".xlsx"));

                return AppKeyHelper.GetKey("EXPORT_FILE_PATH") + fileName + ".xlsx";
            }
        }


        public string BuildExcelSheet(DataTable dtObject, string sheetName, string fileName)
        {

            using (ExcelPackage pck = new ExcelPackage())
            {
                var DT = dtObject;

                //Create the worksheet;
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add(sheetName);

                //Load the datatable into the sheet, starting from cell A1. Print the column names on row 1               

                var dataRange = ws.Cells["A1"].LoadFromDataTable(DT, true);


                var ss = ws.Cells;
                var pp = ss.Take(DT.Columns.Count);

                //Format the header for column 1-3 using (ExcelRange rng = ws.Cells["A1:C1"]) string.Format("{0}:{1}", pp.First(), pp.Last())
                using (ExcelRange rng = ws.Cells[string.Format("{0}:{1}", pp.First(), pp.Last())])
                {
                    rng.Style.Font.Bold = true;
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;                      //Set Pattern for the background to Solid
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));  //Set color to dark blue
                    rng.Style.Font.Color.SetColor(Color.White);
                    rng.AutoFilter = true;
                    rng.AutoFitColumns(45, 150);
                }

                #region PERSISTS FORMAT DATTYPE
                if (DT.Rows.Count > 0)
                {
                    for (int i = 0; i < DT.Columns.Count; i++)
                    {
                        switch (DT.Columns[i].DataType.Name)
                        {

                            case "DateTime": ws.Cells[2, i + 1, dataRange.End.Row, i + 1].Style.Numberformat.Format = "MM-dd-yyyy hh:mm"; break;
                            case "TimeSpan": //ws.Cells[2, i + 1, dataRange.End.Row, i + 1].Style.Numberformat.Format = "hh:mm"; break;
                            case "Decimal":
                            case "Double": //ws.Cells[2, i + 1, dataRange.End.Row, i + 1].Style.Numberformat.Format = "0:0.##"; break;
                            case "SByte":
                            case "Single":
                            case "String":
                            case "Boolean":
                            case "Byte":
                            case "Char":
                            case "Int16":
                            case "Int32":
                            case "Int64":
                            case "UInt16":
                            case "UInt32":
                            case "UInt64": break;//ws.Cells[2, i + 1, dataRange.End.Row, i + 1].Style.Numberformat.Format = "@"; break;
                            default: break;
                        }

                    }
                }

                //ws.Cells[2, 1, dataRange.End.Row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                //ws.Cells[2, 1, dataRange.End.Row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                //ws.Cells[2, 1, dataRange.End.Row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                #endregion

                pck.SaveAs(new FileInfo(AppKeyHelper.GetKey("EXPORT_FILE_PATH") + fileName + ".xlsx"));

                //System.IO.Compression.ZipFile.CreateFromDirectory(@"E:\excell\", fileex, System.IO.Compression.CompressionLevel.Optimal, false);

                return AppKeyHelper.GetKey("EXPORT_FILE_PATH") + fileName + ".xlsx";
            }
        }

        public string BuildExcelSheet<T>(List<T> dtObject, string sheetName, string fileName)
        {
            using (ExcelPackage pck = new ExcelPackage())
            {
                var DT = ListToDataTable(dtObject);

                //Create the worksheet
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add(sheetName);

                //Load the datatable into the sheet, starting from cell A1. Print the column names on row 1
                var dataRange = ws.Cells["A1"].LoadFromDataTable(DT, true);


                var ss = ws.Cells;
                var pp = ss.Take(DT.Columns.Count);

                //Format the header for column 1-3 using (ExcelRange rng = ws.Cells["A1:C1"]) string.Format("{0}:{1}", pp.First(), pp.Last())
                using (ExcelRange rng = ws.Cells[string.Format("{0}:{1}", pp.First(), pp.Last())])
                {
                    rng.Style.Font.Bold = true;
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;                      //Set Pattern for the background to Solid
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));  //Set color to dark blue
                    rng.Style.Font.Color.SetColor(Color.White);
                    rng.AutoFilter = true;
                    rng.AutoFitColumns(45, 150);
                }


                #region PERSISTS FORMAT DATTYPE

                for (int i = 0; i < DT.Columns.Count; i++)
                {
                    switch (DT.Columns[i].DataType.Name)
                    {

                        case "DateTime": ws.Cells[2, i + 1, dataRange.End.Row, i + 1].Style.Numberformat.Format = "MM-dd-yyyy hh:mm"; break;
                        case "TimeSpan": //ws.Cells[2, i + 1, dataRange.End.Row, i + 1].Style.Numberformat.Format = "hh:mm"; break;
                        case "Decimal":
                        case "Double": //ws.Cells[2, i + 1, dataRange.End.Row, i + 1].Style.Numberformat.Format = "0:0.##"; break;
                        case "SByte":
                        case "Single":
                        case "String":
                        case "Boolean":
                        case "Byte":
                        case "Char":
                        case "Int16":
                        case "Int32":
                        case "Int64":
                        case "UInt16":
                        case "UInt32":
                        case "UInt64": break;//ws.Cells[2, i + 1, dataRange.End.Row, i + 1].Style.Numberformat.Format = "@"; break;
                        default: break;
                    }

                }

                //ws.Cells[2, 1, dataRange.End.Row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                //ws.Cells[2, 1, dataRange.End.Row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                //ws.Cells[2, 1, dataRange.End.Row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                #endregion

                pck.SaveAs(new FileInfo(AppKeyHelper.GetKey("EXPORT_FILE_PATH") + fileName + ".xlsx"));
              

                return AppKeyHelper.GetKey("EXPORT_FILE_PATH") + fileName + ".xlsx";
            }
        }
        public static DataTable ListToDataTable<T>(List<T> list)
        {
            DataTable dt = new DataTable();

            foreach (PropertyInfo info in typeof(T).GetProperties())
            {
                dt.Columns.Add(new DataColumn(info.Name, GetNullableType(info.PropertyType)));
            }
            foreach (T t in list)
            {
                DataRow row = dt.NewRow();
                foreach (PropertyInfo info in typeof(T).GetProperties())
                {
                    if (!IsNullableType(info.PropertyType))
                        row[info.Name] = info.GetValue(t, null);
                    else
                        row[info.Name] = (info.GetValue(t, null) ?? DBNull.Value);
                }
                dt.Rows.Add(row);
            }
            return dt;
        }
        private static Type GetNullableType(Type t)
        {
            Type returnType = t;
            if (t.IsGenericType && t.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
            {
                returnType = Nullable.GetUnderlyingType(t);
            }
            return returnType;
        }
        private static bool IsNullableType(Type type)
        {
            return (type == typeof(string) ||
                    type.IsArray ||
                    (type.IsGenericType &&
                     type.GetGenericTypeDefinition().Equals(typeof(Nullable<>))));
        }

        public string BuildMultipleExcelSheets(DataSet dsObject, string sheetName, string fileName)
        {
            using (ExcelPackage pck = new ExcelPackage())
            {
                if (dsObject.Tables.Count > 0)
                {
                    for (int j = 0; j < dsObject.Tables.Count; j++)
                    {

                        var DT = dsObject.Tables[j];

                        //Create the worksheet;
                        ExcelWorksheet ws = pck.Workbook.Worksheets.Add(sheetName + "_" + j);

                        //Load the datatable into the sheet, starting from cell A1. Print the column names on row 1              

                        var dataRange = ws.Cells["A1"].LoadFromDataTable(DT, true);


                        var ss = ws.Cells;
                        var pp = ss.Take(DT.Columns.Count);

                        //Format the header for column 1-3 using (ExcelRange rng = ws.Cells["A1:C1"]) string.Format("{0}:{1}", pp.First(), pp.Last())
                        using (ExcelRange rng = ws.Cells[string.Format("{0}:{1}", pp.First(), pp.Last())])
                        {
                            rng.Style.Font.Bold = true;
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;                      //Set Pattern for the background to Solid
                            rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));  //Set color to dark blue
                            rng.Style.Font.Color.SetColor(Color.White);
                            rng.AutoFilter = true;
                            rng.AutoFitColumns(45, 150);
                        }

                        #region PERSISTS FORMAT DATTYPE
                        if (DT.Rows.Count > 0)
                        {
                            for (int i = 0; i < DT.Columns.Count; i++)
                            {
                                switch (DT.Columns[i].DataType.Name)
                                {

                                    case "DateTime": ws.Cells[2, i + 1, dataRange.End.Row, i + 1].Style.Numberformat.Format = "MM-dd-yyyy hh:mm"; break;
                                    case "TimeSpan":
                                    case "Decimal":
                                    case "Double":
                                    case "SByte":
                                    case "Single":
                                    case "String":
                                    case "Boolean":
                                    case "Byte":
                                    case "Char":
                                    case "Int16":
                                    case "Int32":
                                    case "Int64":
                                    case "UInt16":
                                    case "UInt32":
                                    case "UInt64": break;
                                    default: break;
                                }

                            }
                        }

                        #endregion
                    }
                }
                else
                {
                    DataTable dt = new DataTable();
                    dt.Columns.Add("Info");
                    DataRow r = dt.NewRow();
                    r["Info"] = "No Records Found";
                    dt.Rows.Add(r);
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add(sheetName);

                    var dataRange = ws.Cells["A1"].LoadFromDataTable(dt, true);
                }

                pck.SaveAs(new FileInfo(AppKeyHelper.GetKey("EXPORT_FILE_PATH") + fileName + ".xlsx"));

                return AppKeyHelper.GetKey("EXPORT_FILE_PATH") + fileName + ".xlsx";
            }
        }
        public string BuildMultipleExcelSheetsForPV(DataSet dsObject, string fileName)
        {
            using (ExcelPackage pck = new ExcelPackage())
            {
                if (dsObject.Tables.Count > 0)
                {
                    for (int j = 0; j < dsObject.Tables.Count; j++)
                    {

                        var DT = dsObject.Tables[j];

                        //Create the worksheet;
                        ExcelWorksheet ws = pck.Workbook.Worksheets.Add(dsObject.Tables[j].TableName);

                        //Load the datatable into the sheet, starting from cell A1. Print the column names on row 1              

                        var dataRange = ws.Cells["A1"].LoadFromDataTable(DT, true);
                        
                        var ss = ws.Cells;
                        var pp = ss.Take(DT.Columns.Count);

                        //Format the header for column 1-3 using (ExcelRange rng = ws.Cells["A1:C1"]) string.Format("{0}:{1}", pp.First(), pp.Last())
                        using (ExcelRange rng = ws.Cells[string.Format("{0}:{1}", pp.First(), pp.Last())])
                        {
                            rng.Style.Font.Bold = true;
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;                      //Set Pattern for the background to Solid
                            rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));  //Set color to dark blue
                            rng.Style.Font.Color.SetColor(Color.White);
                            rng.AutoFilter = true;
                            rng.AutoFitColumns(45, 150);
                        }

                        #region PERSISTS FORMAT DATTYPE
                        if (DT.TableName == "Member Demographics" || DT.TableName == "PH-Admission & Discharges" || DT.TableName == "Authorization Class")
                        {
                            if (DT.Rows.Count > 0)
                            {
                                for (int i = 0; i < DT.Rows.Count; i++)
                                {
                                    switch (DT.Rows[i]["Name"].ToString())
                                    {

                                        case "Member Information": 
                                        case "Additional Information":
                                        case "Index": 
                                        case "ADT- Grid Columns": 
                                        case "IP-Authorizations Grid Columns":
                                        case "Care Coordination":
                                        case "Utilization Management":
                                        case "Appeals & Grievances":
                                        case "Population Health":ws.Row(i+2).Style.Font.Bold = true;
                                            //ws.Row(i + 2).Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            //ws.Row(i + 2).Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                                            break;

                                        default: break;
                                    }


                                }
                            }

                            #endregion
                        }
                    }
                }
                else
                {
                    DataTable dt = new DataTable();
                    dt.Columns.Add("Info");
                    DataRow r = dt.NewRow();
                    r["Info"] = "No Records Found";
                    dt.Rows.Add(r);
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Info");

                    var dataRange = ws.Cells["A1"].LoadFromDataTable(dt, true);
                }

                pck.SaveAs(new FileInfo(AppKeyHelper.GetKey("EXPORT_FILE_PATH") + fileName + ".xlsx"));

                return AppKeyHelper.GetKey("EXPORT_FILE_PATH") + fileName + ".xlsx";
            }
        }
        public void ExportExcelByTemplate(DataTable dtObject, string filePath, string fileName)
        {
            FileInfo fiInfo = new FileInfo(filePath);
            if (fiInfo.Exists)
            {
                var DT = dtObject;
                using (ExcelPackage p = new ExcelPackage())
                {
                    using (FileStream stream = new FileStream(fiInfo.FullName, FileMode.Open, FileAccess.ReadWrite))
                    {
                        p.Load(stream);
                        //deleting worksheet if already present in excel file
                        //var wk = p.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Sheet1");
                        //if (wk != null) { p.Workbook.Worksheets.Delete(wk); }
                        ////p.Workbook.CalcMode = 2;
                        //p.Workbook.Worksheets.Add("Sheet1");
                        //p.Workbook.Worksheets.MoveToEnd("Sheet1");
                        //ExcelWorksheet ws = p.Workbook.Worksheets[p.Workbook.Worksheets.Count];
                        ExcelWorksheet ws = p.Workbook.Worksheets[1];
                        var dataRange = ws.Cells["A1"].LoadFromDataTable(DT, true);


                        var ss = ws.Cells;
                        var pp = ss.Take(DT.Columns.Count);

                        //Format the header for column 1-3 using (ExcelRange rng = ws.Cells["A1:C1"]) string.Format("{0}:{1}", pp.First(), pp.Last())
                        using (ExcelRange rng = ws.Cells[string.Format("{0}:{1}", pp.First(), pp.Last())])
                        {
                            rng.Style.Font.Bold = true;
                            rng.Style.Fill.PatternType = ExcelFillStyle.Solid;                      //Set Pattern for the background to Solid
                            rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));  //Set color to dark blue
                            rng.Style.Font.Color.SetColor(Color.White);
                        }
                        //ExcelRange rng = ws.Cells["A2:AI2"];

                        #region PERSISTS FORMAT DATTYPE
                        if (DT.Rows.Count > 0)
                        {
                            for (int i = 0; i < DT.Columns.Count; i++)
                            {
                                switch (DT.Columns[i].DataType.Name)
                                {

                                    case "DateTime": ws.Cells[2, i + 1, dataRange.End.Row, i + 1].Style.Numberformat.Format = "MM-dd-yyyy hh:mm"; break;
                                    case "TimeSpan": //ws.Cells[2, i + 1, dataRange.End.Row, i + 1].Style.Numberformat.Format = "hh:mm"; break;
                                    case "Decimal":
                                    case "Double": //ws.Cells[2, i + 1, dataRange.End.Row, i + 1].Style.Numberformat.Format = "0:0.##"; break;
                                    case "SByte":
                                    case "Single":
                                    case "String":
                                    case "Boolean":
                                    case "Byte":
                                    case "Char":
                                    case "Int16":
                                    case "Int32":
                                    case "Int64":
                                    case "UInt16":
                                    case "UInt32":
                                    case "UInt64": break;//ws.Cells[2, i + 1, dataRange.End.Row, i + 1].Style.Numberformat.Format = "@"; break;
                                    default: break;
                                }

                            }
                        }

                        //ws.Cells[2, 1, dataRange.End.Row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                        //ws.Cells[2, 1, dataRange.End.Row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        //ws.Cells[2, 1, dataRange.End.Row, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                        #endregion
                        //excelWorksheet.SaveAs("E:\\DocRoot\\Temp_Doc\\" +tempFileName);
                        p.SaveAs(new FileInfo(fileName));

                    }
                }
            }
        }

        public DataTable ReadExcelToDataTable(Stream fStream)
        {
            DataTable tbl = new DataTable();
            using (var excel = new ExcelPackage(fStream))
            {
                var ws = excel.Workbook.Worksheets.First();
                var hasHeader = true;  // adjust accordingly
                // add DataColumns to DataTable
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text
                        : String.Format("Column {0}", firstRowCell.Start.Column));

                // add DataRows to DataTable
                int startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = tbl.NewRow();
                    foreach (var cell in wsRow)
                        row[cell.Start.Column - 1] = cell.Text;
                    tbl.Rows.Add(row);
                }
                tbl.TableName = ws.ToString();
            }
            return tbl;
        }

        public DataTable ReadExcelToDataTableBySheet(Stream fStream, string SheetName)
        {
            DataTable tbl = new DataTable();

            using (var excel = new ExcelPackage(fStream))
            {
                var ws = excel.Workbook.Worksheets.First();
                for (int i = 1; i <= excel.Workbook.Worksheets.Count; i++)
                {
                    if (excel.Workbook.Worksheets[i].ToString().Trim() == SheetName)
                    {
                        ws = excel.Workbook.Worksheets[i];
                    }
                }
                // var ws = excel.Workbook.Worksheets.First();
                var hasHeader = true;  // adjust accordingly
                // add DataColumns to DataTable
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text
                        : String.Format("Column {0}", firstRowCell.Start.Column));

                // add DataRows to DataTable
                int startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = tbl.NewRow();
                    foreach (var cell in wsRow)
                        row[cell.Start.Column - 1] = cell.Text;
                    tbl.Rows.Add(row);
                }
                tbl.TableName = ws.ToString();
            }
            return tbl;
        }

        public DataTable SSISReadExcelToDataTable(Stream fStream)
        {
            DataTable tbl = new DataTable();
            using (var excel = new ExcelPackage(fStream))
            {
                var tbbl = excel.Workbook.Worksheets.SelectMany(s => s.Tables);
                if (!tbbl.IsListNullOrEmpty())
                {
                    if (tbbl.First().WorkSheet.ToString().ToUpper() == "HIERARCHY TEMPLATE")
                    {
                        var ws = excel.Workbook.Worksheets["HIERARCHY TEMPLATE"];
                        var hasHeader = true;  // adjust accordingly
                                               // add DataColumns to DataTable
                        foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                            tbl.Columns.Add(hasHeader ? firstRowCell.Text
                                : String.Format("Column {0}", firstRowCell.Start.Column));

                        // add DataRows to DataTable
                        int startRow = hasHeader ? 2 : 1;
                        for (int rowNum = startRow; rowNum <= startRow + 1; rowNum++)
                        {
                            var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                            DataRow row = tbl.NewRow();
                            foreach (var cell in wsRow)
                                row[cell.Start.Column - 1] = cell.Text;
                            tbl.Rows.Add(row);
                        }
                        tbl.TableName = ws.ToString();
                    }
                }
            }
            return tbl;
        }
    }
}