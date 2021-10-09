using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.EditorInput;
using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace CADExtractorLib
{
    class ExcelUtil
    {
        /// <summary>
        /// 엑셀 파일 경로
        /// </summary>
        public string path { get; set; }
        private string extension = ".xlsx";

        private bool IsExist()
        {
            FileInfo fileInfo = new FileInfo(path + extension);

            if (fileInfo.Exists)
            {
                return true;
            }

            return false;

        }

        public void AddData(Dictionary<int, string> layerDict, Dictionary<int, string> areaDict)
        {

            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            ed.WriteMessage("\nWriting Excel....");

            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            app = new Excel.Application();
            bool exist = IsExist();

            if(exist)
            {
                wb = app.Workbooks.Open(path + extension);
                ws = wb.Worksheets.get_Item("Sheet1") as Excel.Worksheet;
            }

            else
            {
                wb = app.Workbooks.Add();
                ws = wb.Worksheets.Add(Type.Missing, wb.Worksheets[1]);
            }
           
            

            try
            {
                for (int i = 0; i < layerDict.Count; i++)
                {
                    ws.Cells[i + 1, 1] = layerDict[i];
                    ws.Cells[i + 1, 2] = areaDict[i];
                }


                ed.WriteMessage("\nComplete");
            }

            catch
            {
                ed.WriteMessage("\nError!");
            }

            finally
            {
                if (exist)
                {
                    wb.Save();
                }

                else
                {
                    wb.SaveAs(path + extension);
                }
                
                wb.Close();
                app.Quit();
            }

        }

    }
}
