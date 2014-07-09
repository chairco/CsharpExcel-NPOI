using System;
using System.Collections.Generic;
using System.Text;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel;
using System.Linq;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;

//NPOI.DLL：NPOI 核心函式庫。
//NPOI.DDF.DLL：NPOI 繪圖區讀寫函式庫。
//NPOI.HPSF.DLL：NPOI 文件摘要資訊讀寫函式庫。
//NPOI.HSSF.DLL：NPOI Excel BIFF 檔案讀寫函式庫。
//NPOI.Util.DLL：NPOI 工具函式庫。
//NPOI.POIFS.DLL：NPOI OLE 格式存取函式庫。
//ICSharpCode.SharpZipLib.DLL：檔案壓縮函式庫。

namespace sample
{

    public class SheetModifyProductList
    {
        public string setProductNumber { get; set; }
    }

    public class SheetHashtable
    {
        public string setdata { get; set;}
    }

    class EXCEL
    {
        public void openexcel(string Datasource)
        {
            IWorkbook wk;
            ISheet st;
            
            using (FileStream fs = new FileStream(Datasource, FileMode.Open, FileAccess.Read))
            {
                if (Datasource.Contains(".xlsx")) //2007
                {
                    wk = new XSSFWorkbook(fs);
                    st = (XSSFSheet)wk.GetSheetAt(0);
                }
                else //2003
                {
                    wk = new HSSFWorkbook(fs);
                    st = (HSSFSheet)wk.GetSheetAt(0);
                }

                int sheetCount = wk.NumberOfSheets;
                List<SheetModifyProductList> modelList = new List<SheetModifyProductList>(); //add

                for (int k = 0; k < sheetCount; k++) //讀出sheetname
                {
                    var hs = wk.GetSheetAt(k); //sheet
                    string sheetname = hs.SheetName; //sheet's name

                    if (sheetname != "N61 FF") continue;

                    var hr = hs.GetRow(0); //row
                    for (int i = hs.FirstRowNum; i <= hs.LastRowNum; i++)
                    {
                        hr = hs.GetRow(i); //column
                        for (int j = hr.FirstCellNum; j < hr.LastCellNum; j++)
                        {
                            SheetModifyProductList model = new SheetModifyProductList(); //add
                            
                            if (hr.GetCell(j) != null)
                            {
                                Console.Write("({0},{1}) = {2} ; ", i, j, hr.GetCell(j).ToString());
                                model.setProductNumber = hr.GetCell(j).ToString() + ";";
                            }
                            modelList.Add(model);
                        }
                        Console.WriteLine("\t\n");
                    }
                }
                wk = null; //全部Sheet讀完關閉Excel
                fs.Close();

                foreach (var item in modelList)
                {
                    Console.Write(item.setProductNumber);
                }
                Console.Read();
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            EXCEL excel = new EXCEL();
            excel.openexcel("C:\\test.xlsx");
        }
    }
}