using Autodesk.AutoCAD.ApplicationServices;
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


        /// <summary>
        /// 엑셀 데이터 저장
        /// layerList와 areaList는 1:1 매칭
        /// </summary>
        /// <param name="layerList">추출한 캐드 레이어 이름</param>
        /// <param name="areaList">polyline으로 닫혀진 면적</param>
        public void AddData(List<string> layerList, List<string> areaList)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            
            ed.WriteMessage("\nWriting Excel....");
            
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {

                app = new Excel.Application();
                wb = app.Workbooks.Open(path);
                ws = wb.Worksheets.get_Item("Sheet1") as Excel.Worksheet;

                for (int i = 0; i < layerList.Count; i++)
                {
                    ws.Cells[i + 1, 1] = layerList[i];
                    ws.Cells[i + 1, 2] = areaList[i];
                }


                ed.WriteMessage("\nComplete");
            }

            catch
            {
                ed.WriteMessage("\nError!");
            }

            finally
            {
                wb.Save();
                wb.Close();
                app.Quit();
            }

        }

    }
}
