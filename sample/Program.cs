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


//NPOI.DLL：NPOI 核心函式庫。
//NPOI.DDF.DLL：NPOI 繪圖區讀寫函式庫。
//NPOI.HPSF.DLL：NPOI 文件摘要資訊讀寫函式庫。
//NPOI.HSSF.DLL：NPOI Excel BIFF 檔案讀寫函式庫。
//NPOI.Util.DLL：NPOI 工具函式庫。
//NPOI.POIFS.DLL：NPOI OLE 格式存取函式庫。
//ICSharpCode.SharpZipLib.DLL：檔案壓縮函式庫。

namespace sample
{
    public class ModelInfo
    {
        public decimal FieldA { get; set; }
        public decimal FieldB { get; set; }
        public decimal FieldC { get; set; }
    }

    public class SheetModifyProductList
    {
        public string setProductNumber { get; set; }
    }

    class EXCEL
    {
        private List<SheetModifyProductList> sheetMappingToList(XSSFSheet sheet)
        {
            List<SheetModifyProductList> modelList = new List<SheetModifyProductList>();
            for (int i = 1; i <= calLastRow(sheet); i++)
            {
                SheetModifyProductList model = new SheetModifyProductList();
                if (sheet.GetRow(i).GetCell(0) != null)
                    //model.setProductNumber(sheet.GetRow(i).GetCell(0).ToString().Trim());
                    model.setProductNumber = sheet.GetRow(i).GetCell(0).ToString().Trim();

                modelList.Add(model);
            }
            return modelList;
        }

        private List<SheetModifyProductList> sheetMappingToList(HSSFSheet sheet)
        {
            List<SheetModifyProductList> modelList = new List<SheetModifyProductList>();
            for (int i = 1; i <= calLastRow(sheet); i++)
            {
                SheetModifyProductList model = new SheetModifyProductList();
                if (sheet.GetRow(i).GetCell(0) != null)
                    //model.setProductNumber(sheet.GetRow(i).GetCell(0).ToString().Trim());
                    model.setProductNumber = sheet.GetRow(i).GetCell(0).ToString().Trim();
                modelList.Add(model);
            }
            return modelList;
        }

        private int calLastRow(XSSFSheet sheet)
        {
            int count = 0;
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                if ((sheet.GetRow(i).GetCell(0) != null || sheet.GetRow(i).GetCell(0).ToString() != "") && (sheet.GetRow(i).GetCell(1) != null || sheet.GetRow(i).GetCell(1).ToString() != "") && (sheet.GetRow(i).GetCell(2) != null || sheet.GetRow(i).GetCell(2).ToString() != "") && (sheet.GetRow(i).GetCell(3) != null || sheet.GetRow(i).GetCell(3).ToString() != "") && (sheet.GetRow(i).GetCell(4) != null || sheet.GetRow(i).GetCell(4).ToString() != "") && (sheet.GetRow(i).GetCell(5) != null || sheet.GetRow(i).GetCell(5).ToString() != ""))
                {
                    count++;
                }
                else
                    break;
            }
            return count;
        }

        private int calLastRow(HSSFSheet sheet)
        {
            int count = 0;
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                if ((sheet.GetRow(i).GetCell(0) != null || sheet.GetRow(i).GetCell(0).ToString() != "") && (sheet.GetRow(i).GetCell(1) != null || sheet.GetRow(i).GetCell(1).ToString() != "") && (sheet.GetRow(i).GetCell(2) != null || sheet.GetRow(i).GetCell(2).ToString() != "") && (sheet.GetRow(i).GetCell(3) != null || sheet.GetRow(i).GetCell(3).ToString() != "") && (sheet.GetRow(i).GetCell(4) != null || sheet.GetRow(i).GetCell(4).ToString() != "") && (sheet.GetRow(i).GetCell(5) != null || sheet.GetRow(i).GetCell(5).ToString() != ""))
                {
                    count++;
                }
                else
                    break;
            }
            return count;
        }
    }

    class Program
    {
        public void openexcel(string Datasource)
        {
            IWorkbook wk;
            ISheet st;
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            
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
                            if (hr.GetCell(j) != null)
                            {
                                //string strcell = hr.GetCell(i) == null ? "0" : hr.GetCell(i).ToString();
                                //Console.Write("({0},{1})={2} ; ", i, j, strcell);
                                Console.Write("({0},{1}) = {2} ; ", i, j, hr.GetCell(j).ToString());
                            }
                        }
                        Console.WriteLine("\t\n");
                    }
                }
                wk = null; //全部Sheet讀完關閉Excel
                fs.Close();
                Console.Read();
            }
        }

        static void Main(string[] args)
        { 
            
        }

        /***
        static void Main(string[] args)
        {
            //new一個用List接的Model
            List<ModelInfo> ModelInfoList = new List<ModelInfo>();

            //會被覆蓋的問題在此，因為只new一次！
            //ModelInfo isModelInfo = new ModelInfo();

            decimal addA = (Decimal)0.1, addB = (Decimal)0.2, addC = (Decimal)0.3;

            //跑三次迴圈
            for (int cnt = 0; cnt <= 3; cnt++)
            {
                ModelInfo isModelInfo = new ModelInfo();
                addA++;
                addB++;
                addC++;

                //每次迴圈Model的每個欄位都增加新的資訊
                isModelInfo.FieldA = addA;
                isModelInfo.FieldB = addB;
                isModelInfo.FieldC = addC;

                //將每次迴圈紀錄的欄位Model，add到List接的Model
                ModelInfoList.Add(isModelInfo);
            }
            
            foreach (var item in ModelInfoList)
            {
                Console.WriteLine(item.FieldA);
                Console.WriteLine(item.FieldB);
                Console.WriteLine(item.FieldC);
            }
            Console.Read();
        }
         *///
    }
}
