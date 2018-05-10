using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Oracle.ManagedDataAccess.Client;

namespace OpenXML
{
    public class Program
    {
        private const string V = @"C:\Users\CAIQICHAO\Desktop\OpenXml.xlsx";

        static void Main(string[] args)
        {
            string strsql = @"select * from container";
            DataTable dt = ConnectData.ExecuteDataTable(strsql, null);
            DataTable[] _dt = new DataTable[1];
            _dt[0] = dt;

            CreateExcel(V, _dt, null);

            Console.WriteLine("Completed!");           
            Console.ReadKey();
        }

        //写入Excel
        static void main()
        {
            var workbook = SpreadsheetDocument.Create(V, SpreadsheetDocumentType.Workbook);
            var workbookPart = workbook.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            workbookPart.Workbook.Sheets = new Sheets();
            var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            sheetPart.Worksheet = new Worksheet();
            Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();//worksheetPart.Worksheet.GetFirstChild<Sheet>();
            string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);
            string sheetName = "Sheet1";
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = 1, Name = sheetName };
            sheets.Append(sheet);
            for (int i = 0; i < 10; i++)
            {
                Row row = new Row();
                for (int j = 0; j < 10; j++)
                {
                    Cell dataCell = new Cell();
                    dataCell.CellValue = new CellValue($"{i + 1}行{j + 1}列");
                    dataCell.DataType = CellValues.String;
                    row.AppendChild(dataCell);
                }
                sheetData.Append(row);
            }
            sheetPart.Worksheet.Append(sheetData);
            workbook.Close();
        }
        //读取Excel
        static void ain()
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(V, false))
            {
                WorkbookPart wbPart = doc.WorkbookPart;
                Sheet mysheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.FirstOrDefault();
                Worksheet worksheet = ((WorksheetPart)wbPart.GetPartById(mysheet.Id)).Worksheet;
                SheetData sheetData = (SheetData)worksheet.ChildElements.FirstOrDefault();

                foreach (var row in sheetData.ChildElements)
                {
                    foreach (var cell in (row as Row).ChildElements)
                    {
                        var cellValue = (cell as Cell).CellValue;
                        if (cellValue != null)
                        {
                            Console.Write(cellValue.Text + " ");
                        }
                    }
                    Console.WriteLine("\n");
                }

            }
        }

        /// <summary>  
        /// 创建excel,并且把dataTable导入到excel中  
        /// </summary>         
        /// <param name="destination">保存路径</param>  
        /// <param name="dataTables">数据源</param>  
        /// <param name="sheetNames">excel中sheet的名称</param>    
        public static void CreateExcel(string destination, DataTable[] dataTables, string[] sheetNames = null)
        {
            using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                workbookPart.Workbook.Sheets = new Sheets();

                uint sheetId = 1;
                bool isAddStyle = false;
                foreach (DataTable table in dataTables)
                {
                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    sheetPart.Worksheet = new Worksheet();
                    if (!isAddStyle)
                    {
                        var stylesPart = workbook.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                        Stylesheet styles = new Stylesheet();
                        styles.Save(stylesPart);
                        isAddStyle = true;
                    }
                    Columns headColumns = CrateColunms(table);
                    sheetPart.Worksheet.Append(headColumns);
                    Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();//worksheetPart.Worksheet.GetFirstChild<Sheet>();
                    string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    if (sheets.Elements<Sheet>().Count() > 0)
                    {
                        sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }
                    string sheetName = string.Empty;
                    if (sheetNames != null)
                    {
                        if (sheetNames.Length >= sheetId)
                        {
                            sheetName = sheetNames[sheetId - 1].ToString();
                        }
                    }
                    else
                    {
                        sheetName = table.TableName ?? sheetId.ToString();
                    }

                    Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                    sheets.Append(sheet);

                    Row headerRow = new Row();

                    List<String> columns = new List<string>();
                    foreach (DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);

                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.StyleIndex = 11;
                        cell.CellValue = new CellValue(column.ColumnName);
                        headerRow.AppendChild(cell);
                    }
                    sheetData.AppendChild(headerRow);

                    foreach (DataRow dsrow in table.Rows)
                    {
                        Row newRow = new Row();
                        foreach (String col in columns)
                        {
                            Cell cell = new Cell
                            {
                                DataType = CellValues.String,
                                StyleIndex = 10,
                                CellValue = new CellValue(dsrow[col].ToString())
                            };
                            newRow.AppendChild(cell);
                        }
                        sheetData.AppendChild(newRow);
                    }
                    sheetPart.Worksheet.Append(sheetData);
                }
                workbook.Close();
            }
        }
        private static Columns CrateColunms(DataTable table)
        {
            int numCols = table.Columns.Count;
            var columns = new Columns();
            for (int col = 0; col < table.Columns.Count; col++)
            {
                int maxWidth = table.Columns[col].ColumnName.Length;
                int valueWidth = 0;
                for (int row = 0; row < table.Rows.Count; row++)
                {
                    valueWidth = table.Rows[row][col].ToString().Trim().Length;
                    if (maxWidth < valueWidth)
                    {
                        maxWidth = valueWidth;
                    }
                }
                Column c = new Column();
                columns.Append(c);
            }
            return columns;
        }

    }

    public class ConnectData
    {
        static string strconn = "User Id=KDMESDB;Password=kdmesdb;Data Source=KDMESDB";

        public static DataTable ExecuteDataTable(string sql,params OracleParameter[] parameters)
        {
            using (OracleConnection conn = new OracleConnection(strconn))
            {
                conn.Open();
                using (OracleCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = sql;
                    cmd.Parameters.AddRange(parameters);
                    OracleDataAdapter dataAdapter = new OracleDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    dataAdapter.Fill(dt);
                    return dt;
                }
            }
        }
    }
}
