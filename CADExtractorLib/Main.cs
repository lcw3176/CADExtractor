using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using CADExtractorLib;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using OpenFileDialog = Autodesk.AutoCAD.Windows.OpenFileDialog;

[assembly: CommandClass(typeof(CADExtractor.Main))]
namespace CADExtractor
{
    public class Main
    {
        /// <summary>
        /// 면적 추출 함수
        /// </summary>
        [CommandMethod("Extract")]
        public void Extract()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = Application.DocumentManager.MdiActiveDocument.Database;

            OpenFileDialog dialog = new OpenFileDialog("Select the Excel Path", "CADExtractor", "xlsx; *;", "CADExtractor", OpenFileDialog.OpenFileDialogFlags.DefaultIsFolder);

            DialogResult sdResult = dialog.ShowDialog();;

            if (sdResult != DialogResult.OK)
            {
                return;
            }

            string excelPath = dialog.Filename;


            /// layerNameResult
            /// 전체 레이어 : *
            /// 여러 개의 레이어 : 쉼표(,) 로 구분
            /// 단일 레이어: 단일 입력
            PromptResult layerNameResult = ed.GetString("\nEnter the Layer Name: ");

            if (layerNameResult.Status != PromptStatus.OK)
            {
                return;
            }

            ExcelUtil excelUtil = new ExcelUtil();
            excelUtil.path = excelPath;

            bool isAllSelected = layerNameResult.StringResult is "*" ? true : false;
            string filteredLayerName = isAllSelected ? "*" : layerNameResult.StringResult.Replace(",", " ");

            List<string> layerList = new List<string>();
            List<double> areaList = new List<double>();

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord currentSpace = trans.GetObject(db.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                BlockTableRecord writeSpace = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                Dictionary<string, int> layerDict = new Dictionary<string, int>();
                DrawOrderTable dot = trans.GetObject(writeSpace.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;

                foreach (ObjectId entId in currentSpace)
                {

                    if (entId.ObjectClass == RXClass.GetClass(typeof(Hatch)))
                    {
                        Hatch hatch = trans.GetObject(entId, OpenMode.ForRead) as Hatch;
                        try
                        {
                            /// 해칭 상태가 아닌데 해치로 잡혀서 면적 가져오다가 에러 잡히는 경우가 있다.

                            if (string.IsNullOrEmpty(hatch.Area.ToString()))
                            {
                                continue;
                            }

                            if (filteredLayerName.Contains(hatch.Layer) || isAllSelected)
                            {

                                if (hatch.GetLoopAt(0).IsPolyline)
                                {
                                    DBText text = new DBText();
                                    text.SetDatabaseDefaults();

                                    string dTextContent = hatch.Layer.Contains("-") ? hatch.Layer.Split('-')[1].Substring(0, 2): hatch.Layer.Substring(0, 2);

                                    if (!layerDict.ContainsKey(hatch.Layer))
                                    {
                                        layerDict.Add(hatch.Layer, 1);
                                    }

                                    text.TextString = string.Format("{0}{1} {2}", dTextContent, layerDict[hatch.Layer]++, Math.Round(hatch.Area).ToString());
                                    text.Height = 10;

                                    BulgeVertexCollection bulges = hatch.GetLoopAt(0).Polyline;
                                    double xTemp = 0;
                                    double yTemp = 0;

                                    for (int j = 0; j < bulges.Count; j++)
                                    {
                                        BulgeVertex bulg = bulges[j];
                                        Point2d pt = bulg.Vertex;
                                        xTemp += pt.X;
                                        yTemp += pt.Y;
                                    }


                                    text.Position = new Point3d(xTemp / bulges.Count, yTemp / bulges.Count, 0);
                                    writeSpace.AppendEntity(text);
                                    dot.MoveToBottom(new ObjectIdCollection() { entId });
                                    trans.AddNewlyCreatedDBObject(text, true);
                                }

                                layerList.Add(hatch.Layer);
                                areaList.Add(hatch.Area);
                            }

                        }

                        catch
                        {
                            continue;
                        }
                        
                        
                    }
                }

                trans.Commit();
            }


            excelUtil.AddData(layerList, areaList);

        }


        /// <summary>
        /// 면적 합친 후 계산하는 함수
        /// </summary>
        [CommandMethod("Merge")]
        public void Merge()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = Application.DocumentManager.MdiActiveDocument.Database;


            PromptResult layerNameResult = ed.GetString("\nEnter Layer Number: ");

            if (layerNameResult.Status != PromptStatus.OK)
            {
                return;
            }


            PromptResult layerNumberResult = ed.GetString("\nEnter Layer Number: ");

            if (layerNumberResult.Status != PromptStatus.OK)
            {
                return;
            }

            string layerName = layerNameResult.StringResult;
            List<int> layerNumberName = new List<int>();

            foreach (string i in layerNumberResult.StringResult.ToString().Split(','))
            {
                ed.WriteMessage(i);
                if (string.IsNullOrEmpty(i))
                {
                    continue;
                }

                if (i.Contains("~"))
                {
                    int startNumber = 0;
                    int endNumber = 0;

                    startNumber = int.Parse(i.Split('~')[0]);
                    endNumber = int.Parse(i.Split('~')[1]);

                    for(int j = startNumber; j <= endNumber; j++)
                    {
                        layerNumberName.Add(j);
                    }
                }

                else
                {
                    layerNumberName.Add(int.Parse(i));
                }
            }
            

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord currentSpace = trans.GetObject(db.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                BlockTableRecord writeSpace = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                Dictionary<string, int> layerDict = new Dictionary<string, int>();
                DrawOrderTable dot = trans.GetObject(writeSpace.DrawOrderTableId, OpenMode.ForWrite) as DrawOrderTable;

                DBText text = new DBText();
                text.SetDatabaseDefaults();


                double hatchArea = 0;
                double xTemp = 0;
                double yTemp = 0;
                int count = 0;

                foreach (ObjectId entId in currentSpace)
                {

                    if (entId.ObjectClass == RXClass.GetClass(typeof(Hatch)))
                    {
                        Hatch hatch = trans.GetObject(entId, OpenMode.ForRead) as Hatch;
                        try
                        {
                            /// 해칭 상태가 아닌데 해치로 잡혀서 면적 가져오다가 에러 잡히는 경우가 있다.

                            if (string.IsNullOrEmpty(hatch.Area.ToString()))
                            {
                                continue;
                            }

                            bool result = false;

                            foreach(int i in layerNumberName)
                            {
                                if (hatch.Layer.Contains(i.ToString()))
                                {
                                    result = true;
                                }
                            }


                            if (layerName.Contains(hatch.Layer) && result)
                            {


                                if (hatch.GetLoopAt(0).IsPolyline)
                                {
                                 
                                    BulgeVertexCollection bulges = hatch.GetLoopAt(0).Polyline;
     

                                    for (int j = 0; j < bulges.Count; j++)
                                    {
                                        BulgeVertex bulg = bulges[j];
                                        Point2d pt = bulg.Vertex;
                                        xTemp += pt.X;
                                        yTemp += pt.Y;
                                    }

                                    count += bulges.Count;
                                }
                            }

                            hatchArea += hatch.Area;

                        }

                        catch
                        {
                            continue;
                        }


                    }


                  
                }

                text.TextString = string.Format("{0}-{1}", "병합", Math.Round(hatchArea).ToString());
                text.Height = 70;


                text.Position = new Point3d(0, 0, 0);
                writeSpace.AppendEntity(text);
                trans.AddNewlyCreatedDBObject(text, true);

                trans.Commit();
            }



        }

    }
}