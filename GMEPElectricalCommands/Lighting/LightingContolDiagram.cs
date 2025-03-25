using System;
using System.Diagnostics.Metrics;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Accord.Statistics.Distributions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using ElectricalCommands.ElectricalEntity;
using GMEPElectricalCommands.GmepDatabase;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Collections.Generic;

namespace ElectricalCommands.Lighting
{
  class LightingContolDiagram
  {
    LightingTimeClock TimeClock;
    public LightingContolDiagram(LightingTimeClock timeClock) {
      this.TimeClock = timeClock;
    }
      
    public void InitializeDiagramBase() {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      GmepDatabase gmepDb = new GmepDatabase();
      string projectId = gmepDb.GetProjectId(CADObjectCommands.GetProjectNoFromFileName());
  
      Point3d point;
      using (Transaction tr = db.TransactionManager.StartTransaction()) {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord baseBlock = tr.GetObject(bt["LTG CTRL BASE"], OpenMode.ForWrite) as BlockTableRecord;
        BlockJig blockJig = new BlockJig();
        PromptResult res = blockJig.DragMe(baseBlock.ObjectId, out point);
        if (res.Status == PromptStatus.OK) {
          BlockTableRecord curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
          BlockReference br = new BlockReference(point, baseBlock.ObjectId);
          curSpace.AppendEntity(br);
          tr.AddNewlyCreatedDBObject(br, true);
        }
        tr.Commit();
      }

    }
  }
}
