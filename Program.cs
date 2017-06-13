using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using NPOI;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Data;
using NPOI.SS.UserModel;
using System.Configuration;


namespace DecodeTool
{
    class DecodeColumn
    {
        public string ColumnID { get; set; }
    }


    class Program
    {
        static void Main(string[] args)
        {
            //File Input
            Console.Write("請將Excel檔案拖曳至此視窗：");
            string filePath = Console.ReadLine().Trim('"');

            //Import Data 
            var data = ReadExcelFile(filePath);

            //Export Data
            ExportExcel(filePath, data);

            Console.WriteLine("FINISH!");

            //DEBUG PAUSE
            Console.ReadLine();
        }

        static DataTable ReadExcelFile(string filePath)
        {
            HSSFWorkbook hssfworkbook = null;
            DataTable data = new DataTable();
            var decodeColumn = ConfigurationSettings.AppSettings["DecodeColumn"].Split(';').ToList<string>();
            try
            {
                System.IO.FileStream fs = new FileStream(filePath, FileMode.Open);
                hssfworkbook = new HSSFWorkbook(fs);
                fs.Close();
            }
            catch
            {
                Console.WriteLine(@"檔案開啟異常，請確認上傳檔案副檔名為""xls""的Excel檔案");
                return data;
            }

            HSSFSheet sheet = (HSSFSheet)hssfworkbook.GetSheetAt(0);

            //loop data
            int rowCount = sheet.LastRowNum;
            int RowStart = sheet.FirstRowNum;
            var decodeColumnID = new List<int>();

            for (int i = RowStart; i <= sheet.LastRowNum; i++)
            {
                HSSFRow row = (HSSFRow)sheet.GetRow(i);
                DataRow workRow = data.NewRow();
                //標題列處理               
                for (int j = 0; j < row.Cells.Count(); j++)
                {

                    if (i == 0)
                    {
                        DataColumn column;
                        column = new DataColumn();
                        column.ColumnName = row.GetCell(j).ToString();
                        //find the tagert column
                        if (decodeColumn.Exists(x => x == row.GetCell(j).ToString()))
                        {
                            decodeColumnID.Add(j);
                        }
                        column.DataType = System.Type.GetType("System.String");
                        column.Unique = false;
                        column.AutoIncrement = false;
                        column.Caption = row.GetCell(j).ToString();
                        column.ReadOnly = false;
                        data.Columns.Add(column);
                        continue;
                    }
                    else
                    {
                        if (row.GetCell(j) == null)
                        { workRow[j] = string.Empty; }
                        else
                        {
                            if (decodeColumnID.Exists(x => x == j))
                            {
                                workRow[j] = StringDecode(row.GetCell(j).ToString()).Trim();
                            }
                            else
                            {
                                if (row.GetCell(j).ToString().IndexOf("&#") > -1)
                                {
                                    workRow[j] = UnicodeDecode(row.GetCell(j).ToString()).Trim();
                                }
                                else
                                {
                                    workRow[j] = row.GetCell(j);
                                }
                            }
                        }
                    }
                }
                if (i != 0)
                {
                    data.Rows.Add(workRow);
                }
            }
            return data;
        }



        static string UnicodeDecode(string targetStr)
        {
            string result = HttpUtility.HtmlDecode(targetStr);
            return result;
        }

        static string StringDecode(string targetStr)
        {
            if (string.IsNullOrEmpty(targetStr))
            {
                return string.Empty;
            }
            string result = string.Empty;
            //result = Decode(targetStr);  //DECODE FUNCTION IN HERE!!!
            if (result.IndexOf("&#") > -1)
            {
                result = UnicodeDecode(result);
            }
            return result;
        }


        static void ExportExcel(string path, DataTable data)
        {
            IWorkbook wb = new HSSFWorkbook();
            ISheet ws;
            ////建立Excel 2007檔案   
            if (data.TableName != string.Empty)
            {
                ws = wb.CreateSheet(data.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < data.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(data.Columns[i].ColumnName);
            }

            for (int i = 0; i < data.Rows.Count; i++)
            {
                ws.CreateRow(i + 1);
                for (int j = 0; j < data.Columns.Count; j++)
                {
                    ws.GetRow(i + 1).CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                }
            }

            path = path.Insert(path.LastIndexOf('.'), "_result"); //同個位置檔名後面+_result後存放
            FileStream file = new FileStream(path, FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();
        }


    }
}
