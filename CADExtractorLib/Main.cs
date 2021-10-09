using Autodesk.AutoCAD.ApplicationServices.Core;
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
        [CommandMethod("Extract")]
        public void Extract()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = Application.DocumentManager.MdiActiveDocument.Database;
            

            var excelPathResult = ed.GetString("\nEnter the Excel Path: ");

            if (excelPathResult.Status != PromptStatus.OK)
            {
                return;
            }

            var layerNameResult = ed.GetString("\nEnter the Layer Name: ");

            if (layerNameResult.Status != PromptStatus.OK)
            {
                return;
            }

            ExcelUtil excelUtil = new ExcelUtil();
            excelUtil.path = excelPathResult.StringResult;

            Dictionary<int, string> layerDict = new Dictionary<int, string>();
            Dictionary<int, string> areaDict = new Dictionary<int, string>();

            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord currentSpace = trans.GetObject(db.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;

                int index = 0;

                foreach (ObjectId entId in currentSpace)
                {
                    if (entId.ObjectClass == RXClass.GetClass(typeof(Polyline)))
                    {
                        Polyline pline = trans.GetObject(entId, OpenMode.ForRead) as Polyline;

                        if (pline.Closed && layerNameResult.StringResult == pline.Layer)
                        {
                            layerDict.Add(index, pline.Layer);
                            areaDict.Add(index, Convert.ToInt32(pline.Area).ToString());
                            index++;
                        }
                    }
                }
            }

            excelUtil.AddData(layerDict, areaDict);

        }
    
    }
}