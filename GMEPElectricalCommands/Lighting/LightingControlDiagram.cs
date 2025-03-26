﻿using System;
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
    Point3d InteriorPosition;
    Point3d ExteriorPosition;
    List<ElectricalEntity.LightingLocation> Locations;
    List<ElectricalEntity.LightingFixture> Fixtures;
    public LightingControlDiagram(LightingTimeClock timeClock) {
      this.TimeClock = timeClock;
      this.InitializeDiagram();
    }

    public void InitializeDiagram() {
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
                if (attDef.Tag == "VOLTAGE") {
                  attRef.TextString = TimeClock.Voltage + "V";
                }
                if (attDef.Tag == "SWITCH") {
                  attRef.TextString = TimeClock.BypassSwitchName;
                }
                if (attDef.Tag == "LOCATION") {
                  attRef.TextString = TimeClock.BypassSwitchLocation;
                }
                if (attDef.Tag == "DESCRIPTION") {
                  attRef.TextString = "("+TimeClock.Voltage+"V) ASTRONOMICAL 365-DAY PROGRAMMABLE TIME SWITCH \""+TimeClock.Name+"\" LOCATED ADJACENT TO PANEL \""+panelName+"\" IN A NEMA 1 ENCLOSURE. (TORK #ELC74 OR APPROVED EQUAL)\r\n";
                }
                br.AttributeCollection.AppendAttribute(attRef);
                tr.AddNewlyCreatedDBObject(attRef, true);
              }
            }
          }

          double x = 0;
          double y = 0;
          double x2 = 0;
          double y2 = 0;
          foreach (DynamicBlockReferenceProperty property in br.DynamicBlockReferencePropertyCollection) {
            if (property.PropertyName == "Position1 X") {
             ed.WriteMessage(property.Value.ToString());
              x = (double)property.Value;
            }
            if (property.PropertyName == "Position1 Y") {
              ed.WriteMessage(property.Value.ToString());
              y = (double)property.Value;
            }
            if (property.PropertyName == "Position2 X") {
              ed.WriteMessage(property.Value.ToString());
              x2 = (double)property.Value;
            }
            if (property.PropertyName == "Position2 Y") {
              ed.WriteMessage(property.Value.ToString());
              y2 = (double)property.Value;
            }      
          }
          InteriorPosition = new Point3d(point.X + x, point.Y + y, 0);
          ExteriorPosition = new Point3d(point.X + x2, point.Y + y2, 0);
          

          tr.Commit();
        }
      }

    }
  }
}
