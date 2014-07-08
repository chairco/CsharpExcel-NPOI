using System;
using System.Collections.Generic;
using System.Text;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel;
using System.Linq;
using System.Collections.Generic;
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
    class Program
    {
        public void _openexcel(string Datasource)
        {
            using (FileStream fs = new FileStream(Datasource, FileMode.Open, FileAccess.Read))
            {
                IWorkbook wb;

                if (Datasource.Contains(".xlsx"))
                    wb = new XSSFWorkbook(fs);
                else
                    wb = new HSSFWorkbook(fs);
            }
        }

        public void openexcel(string Datasource)
        {
            XSSFWorkbook wk2007;
            XSSFSheet st2007;
            HSSFWorkbook wk2003;
            HSSFSheet st2003;
            
            using (FileStream fs = new FileStream(Datasource, FileMode.Open, FileAccess.Read))
            {
                if (Datasource.Contains(".xlsx")) //2007
                {
                    wk2007 = new XSSFWorkbook(fs);
                    st2007 = (XSSFSheet)wk2007.GetSheetAt(0);
                }
                else //2003
                {
                    wk2003 = new HSSFWorkbook(fs);
                    st2003 = (HSSFSheet)wk2003.GetSheetAt(0);
                }

                fs.Close();
            }
        }

        private List<SheetModifyProductList> sheetMappingToList(XSSFSheet sheet)
        {
            List<SheetModifyProductList> modelList = new List<SheetModifyProductList>();
            for (int i = 1; i <= calLastRow(sheet); i++)
            {
                SheetModifyProductList model = new SheetModifyProductList();
                if (sheet.GetRow(i).GetCell(0) != null)
                    model.setProductNumber(sheet.GetRow(i).GetCell(0).ToString().Trim());

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
                    model.setProductNumber(sheet.GetRow(i).GetCell(0).ToString().Trim());

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

        static void Main(string[] args)
        {
            
        }
    }
}
