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
    
    }
}