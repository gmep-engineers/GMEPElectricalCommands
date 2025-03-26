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
using Editor = Autodesk.AutoCAD.EditorInput.Editor;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Bibliography;
using Autodesk.AutoCAD.GraphicsInterface;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Drawing;

namespace ElectricalCommands.Lighting {
  class LightingControlDiagram {
    LightingTimeClock TimeClock;
    public LightingControlDiagram(LightingTimeClock timeClock) {
      this.TimeClock = timeClock;
    }

    public void InitializeDiagramBase() {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      GmepDatabase gmepDb = new GmepDatabase();
      string projectId = gmepDb.GetProjectId(CADObjectCommands.GetProjectNoFromFileName());
      List<ElectricalEntity.Panel> panels = gmepDb.GetPanels(projectId);
      ElectricalEntity.Panel panel = panels.Find(p => p.Id == TimeClock.AdjacentPanelId);
      string panelName = panel.Name;

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

          foreach (ObjectId objId in baseBlock) {
            DBObject obj = tr.GetObject(objId, OpenMode.ForRead);
            AttributeDefinition attDef = obj as AttributeDefinition;
            if (attDef != null && !attDef.Constant) {
              using (AttributeReference attRef = new AttributeReference()) {
                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                attRef.Position = attDef.Position.TransformBy(br.BlockTransform);
                if (attDef.Tag == "NAME") {
                  attRef.TextString = TimeClock.Name;
                }
                if (attDef.Tag == "VOLTAGE") {
                  attRef.TextString = TimeClock.Voltage;
                }
                if (attDef.Tag == "SWITCH") {
                  attRef.TextString = TimeClock.BypassSwitchName;
                }
                if (attDef.Tag == "LOCATION") {
                  attRef.TextString = TimeClock.BypassSwitchLocation;
                }
                if (attDef.Tag == "PANEL") {
                  attRef.TextString = panelName;
                }
                br.AttributeCollection.AppendAttribute(attRef);
                tr.AddNewlyCreatedDBObject(attRef, true);
              }
            }
          }
          br.ResetBlock();

          tr.Commit();
        }
      }

    }
  }
}
