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
        public void AddData(List<string> layerList, List<double> areaList)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            
            ed.WriteMessage("\nWriting Excel....");
            
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                Dictionary<string, int> layerColumnDict = new Dictionary<string, int>();
                Dictionary<string, int> layerRowDict = new Dictionary<string, int>();
                Dictionary<string, double> areaValueDict = new Dictionary<string, double>();
                
                app = new Excel.Application();
                wb = app.Workbooks.Open(path);
                ws = wb.Worksheets.get_Item("Sheet1") as Excel.Worksheet;

                int index = 1 + 1;

                // 레이어 목록 삽입
                foreach(string i in layerList)
                {
                    if (!layerColumnDict.ContainsKey(i))
                    {
                        layerColumnDict.Add(i, index);
                        layerRowDict.Add(i, 1);
                        areaValueDict.Add(i, 0);

                        ws.Cells[1, index++] = i;
                    }

                    continue;

                }

                int rowMax = 0;

                // 면적 삽입
                for (int i = 0; i < layerList.Count; i++)
                {
                    ws.Cells[++layerRowDict[layerList[i]], layerColumnDict[layerList[i]]] = areaList[i];

                    if(layerRowDict[layerList[i]] > rowMax)
                    {
                        rowMax = layerRowDict[layerList[i]];
                    }

                    areaValueDict[layerList[i]] += areaList[i];
                }

                // 엑셀 제일 왼쪽에 면적 인덱스 삽입
                for(int i = 0; i < rowMax - 1; i++)
                {
                    ws.Cells[i + 2, 1] = i + 1;
                }


                int rowIndex = 1;
                int columnIndex = layerColumnDict.Count + 1;

                foreach(string i in layerColumnDict.Keys)
                {
                    ws.Cells[rowIndex, columnIndex + 2] = i;
                    ws.Cells[rowIndex, columnIndex + 3] = areaValueDict[i];
                    ws.Cells[rowIndex++, columnIndex + 4] = Math.Round(areaValueDict[i], MidpointRounding.AwayFromZero).ToString();
                }



                ed.WriteMessage("\nComplete");
            }

            catch(Exception ex)
            {
                
                ed.WriteMessage("\nError!");
                ed.WriteMessage(ex.ToString());
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
