using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using CADExtractorLib;
using System;
using System.Collections.Generic;

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
            

            PromptResult excelPathResult = ed.GetString("\nEnter the Excel Path: ");

            if (excelPathResult.Status != PromptStatus.OK)
            {
                return;
            }

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
            excelUtil.path = excelPathResult.StringResult;

            bool isAllSelected = layerNameResult.StringResult is "*" ? true : false;
            string filteredLayerName = isAllSelected ? "*" : layerNameResult.StringResult.Replace(",", " ");

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

                        if (pline.Closed && (filteredLayerName.Contains(pline.Layer) || isAllSelected))
                        {
                            layerList.Add(pline.Layer);
                            areaList.Add(Convert.ToInt32(pline.Area).ToString());
                        }
                    }
                }
            }

            excelUtil.AddData(layerList, areaList);

        }
    
    }
}