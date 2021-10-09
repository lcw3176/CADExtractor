using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


[assembly: CommandClass(typeof(CADExtractor.Main))]
namespace CADExtractor
{
    public class Main
    {
        [CommandMethod("Extract")]
        public void Extract()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = Application.DocumentManager.MdiActiveDocument.Database;
            
            var promptResult = ed.GetString("\nEnter the Excel Path: ");

            if (promptResult.Status != PromptStatus.OK)
            {
                return;
            }
            List<string> layerList = new List<string>();
            List<string> areaList = new List<string>();

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord currentSpace = trans.GetObject(db.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                

                foreach (ObjectId entId in currentSpace)
                {
                    if (entId.ObjectClass == RXClass.GetClass(typeof(Polyline)))
                    {
                        Polyline pline = trans.GetObject(entId, OpenMode.ForRead) as Polyline;

                        if (pline.Closed)
                        {
                            layerList.Add(pline.Layer);
                            areaList.Add(Convert.ToInt32(pline.Area).ToString());
                        }
                    }
                }
            }


            AddData(layerList, areaList);

        }



        private bool AddData(List<string> layerName, List<string> area)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            ed.WriteMessage("\nWriting Excel....");

            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            app = new Excel.Application();
            wb = app.Workbooks.Open("D:/test" + ".xlsx");
            ws = wb.Worksheets.get_Item("Sheet1") as Excel.Worksheet;

            try
            {
                for (int i = 0; i < layerName.Count; i++)
                {
                    ws.Cells[i + 1, 1] = layerName[i];
                    ws.Cells[i + 1, 2] = area[i];
                }


                ed.WriteMessage("\nComplete");
                return true;
            }

            catch
            {

                ed.WriteMessage("\nError!");
                return false;
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