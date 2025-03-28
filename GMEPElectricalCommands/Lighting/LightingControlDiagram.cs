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
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;


namespace ElectricalCommands.Lighting {
  class LightingControlDiagram {
    LightingTimeClock TimeClock;
    Point3d InteriorPosition;
    Point3d ExteriorPosition;
    List<ElectricalEntity.LightingLocation> Locations;
    List<ElectricalEntity.LightingFixture> Fixtures;
    public LightingControlDiagram(LightingTimeClock timeClock, List<ElectricalEntity.LightingLocation> locations, List<ElectricalEntity.LightingFixture> fixtures) {
      this.TimeClock = timeClock;
      this.Fixtures = fixtures;
      this.Locations = locations;
      this.CreateDiagram();
    }
    public void CreateDiagram() {
      InitializeDiagram();
      GraphInteriorLighting();
      GraphExteriorLighting();
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
                  attRef.TextString = "(" + TimeClock.Voltage + "V) ASTRONOMICAL 365-DAY PROGRAMMABLE TIME SWITCH \"" + TimeClock.Name + "\" LOCATED ADJACENT TO PANEL \"" + panelName + "\" IN A NEMA 1 ENCLOSURE. (TORK #ELC74 OR APPROVED EQUAL)\r\n";
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
    public void GraphInteriorLighting() {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      List<LightingLocation> indoorLocations = Locations.Where(location => !location.Outdoor).ToList();

      using (Transaction tr = db.TransactionManager.StartTransaction()) {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
        bool isFirstFlag = true;
        foreach (LightingLocation location in indoorLocations) {
          if (!isFirstFlag) {
            try {
              Point3d startPoint = InteriorPosition;
              Point3d endPoint = new Point3d(startPoint.X, startPoint.Y - 1.75, startPoint.Z);
              Line verticalLine = new Line(startPoint, endPoint);
              curSpace.AppendEntity(verticalLine);
              tr.AddNewlyCreatedDBObject(verticalLine, true);

              Circle circle = new Circle(endPoint, Vector3d.ZAxis, .020);
              curSpace.AppendEntity(circle);
              tr.AddNewlyCreatedDBObject(circle, true);

              Hatch hatch = new Hatch();
              curSpace.AppendEntity(hatch);
              tr.AddNewlyCreatedDBObject(hatch, true);

              hatch.SetDatabaseDefaults();
              hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
              hatch.Associative = true;
              hatch.AppendLoop(HatchLoopTypes.Default, new ObjectIdCollection { circle.ObjectId });
              hatch.EvaluateHatch(true);

              InteriorPosition = endPoint;
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex) {
              ed.WriteMessage($"\nError: {ex.Message}");
            }
          }
          isFirstFlag = false;

          GraphInteriorLocationSection(location);

        }

        tr.Commit();
      }
    }
    public void GraphExteriorLighting() {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      List<LightingLocation> indoorLocations = Locations.Where(location => location.Outdoor).ToList();

      using (Transaction tr = db.TransactionManager.StartTransaction()) {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

        bool isFirstFlag = true;
        foreach (LightingLocation location in indoorLocations) { 
          if (!isFirstFlag) {
            try {
              Point3d startPoint = ExteriorPosition;
              Point3d endPoint = new Point3d(startPoint.X, startPoint.Y - 1.75, startPoint.Z);
              Line verticalLine = new Line(startPoint, endPoint);
              curSpace.AppendEntity(verticalLine);
              tr.AddNewlyCreatedDBObject(verticalLine, true);

              Circle circle = new Circle(endPoint, Vector3d.ZAxis, .020);
              curSpace.AppendEntity(circle);
              tr.AddNewlyCreatedDBObject(circle, true);

              Hatch hatch = new Hatch();
              curSpace.AppendEntity(hatch);
              tr.AddNewlyCreatedDBObject(hatch, true);

              hatch.SetDatabaseDefaults();
              hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
              hatch.Associative = true;
              hatch.AppendLoop(HatchLoopTypes.Default, new ObjectIdCollection { circle.ObjectId });
              hatch.EvaluateHatch(true);

              ExteriorPosition = endPoint;
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex) {
              ed.WriteMessage($"\nError: {ex.Message}");
            }
          }
          isFirstFlag = false;

          //GraphExteriorLocationSection(location)

        }

        tr.Commit();
      }
    }
    private void GraphInteriorLocationSection(LightingLocation location) {

      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      List<LightingFixture> fixturesAtLocation = Fixtures.Where(fixture => fixture.LocationId == location.Id).ToList();

      List<LightingFixture> uniqueFixtures = new List<LightingFixture>();
      var seenCombinations = new HashSet<(string ParentName, int Circuit)>();
      foreach (var fixture in fixturesAtLocation) {
        var combination = (fixture.ParentName, fixture.Circuit);
        if (!seenCombinations.Contains(combination)) {
          seenCombinations.Add(combination);
          uniqueFixtures.Add(fixture);
        }
      }

      //This method will graph the section for each interior lighting location
      using (Transaction tr = db.TransactionManager.StartTransaction()) {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
        TextStyleTable textStyleTable = tr.GetObject(db.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
        ObjectId textStyleId = textStyleTable["Standard"];

        //Graphing Conduit Loop
        Point3d startPoint = InteriorPosition;
        Point3d endPoint = new Point3d(startPoint.X + 1, startPoint.Y, startPoint.Z);
        Line horizontalLine = new Line(startPoint, endPoint);
        curSpace.AppendEntity(horizontalLine);
        tr.AddNewlyCreatedDBObject(horizontalLine, true);

        startPoint = endPoint;
        endPoint = new Point3d(endPoint.X, endPoint.Y - .1, endPoint.Z);
        Line verticalLine = new Line(startPoint, endPoint);
        curSpace.AppendEntity(verticalLine);
        tr.AddNewlyCreatedDBObject(verticalLine, true);

        double radius = .09;
        Point3d circleCenter = new Point3d(endPoint.X, endPoint.Y - radius, endPoint.Z);
        Circle circle = new Circle(circleCenter, Vector3d.ZAxis, radius);
        circle.Layer = "E-TEXT";
        curSpace.AppendEntity(circle); 
        tr.AddNewlyCreatedDBObject(circle, true);

        DBText text = new DBText();
        text.Position = circleCenter;
        text.Height = radius;
        text.TextString = "C";
        text.HorizontalMode = TextHorizontalMode.TextCenter;
        text.VerticalMode = TextVerticalMode.TextVerticalMid;
        text.AlignmentPoint = circleCenter;
        text.TextStyleId = textStyleId;
        text.Justify = AttachmentPoint.MiddleCenter;
        text.Layer = "E-TEXT";
        curSpace.AppendEntity(text);
        tr.AddNewlyCreatedDBObject(text, true);


        startPoint = new Point3d(endPoint.X, endPoint.Y - radius*2, endPoint.Z);
        endPoint = new Point3d(startPoint.X, startPoint.Y -.1, startPoint.Z);
        Line verticalLine2 = new Line(startPoint, endPoint);
        curSpace.AppendEntity(verticalLine2);
        tr.AddNewlyCreatedDBObject(verticalLine2, true);

        startPoint = endPoint;
        endPoint = new Point3d(endPoint.X - .4, endPoint.Y, endPoint.Z);
        Line horizontalLine2 = new Line(startPoint, endPoint);
        curSpace.AppendEntity(horizontalLine2);
        tr.AddNewlyCreatedDBObject(horizontalLine2, true);

        Point3d textCenter = new Point3d(endPoint.X - .06, endPoint.Y, endPoint.Z);

        DBText text2 = new DBText();
        text2.Position = textCenter;
        text2.Height = radius;
        text2.TextString = "N";
        text2.HorizontalMode = TextHorizontalMode.TextCenter;
        text2.VerticalMode = TextVerticalMode.TextVerticalMid;
        text2.AlignmentPoint = textCenter;
        text2.TextStyleId = textStyleId;
        text2.Justify = AttachmentPoint.MiddleCenter;
        text2.Layer = "E-TEXT";
        curSpace.AppendEntity(text2);
        tr.AddNewlyCreatedDBObject(text2, true);

        //Graph Circuits
        Point3d arrowPosition = new Point3d(InteriorPosition.X + 1.3, InteriorPosition.Y + .9, endPoint.Z);
        foreach (LightingFixture fixture in uniqueFixtures) {
          //Begin Arrow
          startPoint = arrowPosition;
          endPoint = new Point3d(startPoint.X, startPoint.Y - 1.06, startPoint.Z);
          Line beginArrow = new Line(startPoint, endPoint);
          curSpace.AppendEntity(beginArrow);
          tr.AddNewlyCreatedDBObject(beginArrow, true);

          //Panel & Circuit Label
          DBText label = new DBText();
          label.Position = new Point3d(startPoint.X, startPoint.Y, startPoint.Z);
          label.Rotation = (Math.PI / 2);
          label.Height = radius * .9;
          label.TextString = fixture.ParentName + "-" + fixture.Circuit.ToString();
          label.HorizontalMode = TextHorizontalMode.TextCenter;
          label.VerticalMode = TextVerticalMode.TextVerticalMid;
          label.AlignmentPoint = new Point3d(startPoint.X, startPoint.Y, startPoint.Z);
          label.Justify = AttachmentPoint.BottomRight;
          label.Layer = "E-TEXT";
          curSpace.AppendEntity(label);
          tr.AddNewlyCreatedDBObject(label, true);

          //Draw Horizontal lines
          Line separator = new Line(new Point3d(endPoint.X - .07, endPoint.Y, endPoint.Z), new Point3d(endPoint.X + .07, endPoint.Y, endPoint.Z));
          separator.Layer = "E-TEXT";
          curSpace.AppendEntity(separator);
          tr.AddNewlyCreatedDBObject(separator, true);

          startPoint = new Point3d(endPoint.X, endPoint.Y - .05, endPoint.Z);
          endPoint = new Point3d(startPoint.X, startPoint.Y -.4, startPoint.Z);

          Line separator2 = new Line(new Point3d(startPoint.X - .07, startPoint.Y, startPoint.Z), new Point3d(startPoint.X + .07, startPoint.Y, startPoint.Z));
          separator2.Layer = "E-TEXT";
          curSpace.AppendEntity(separator2);
          tr.AddNewlyCreatedDBObject(separator2, true);


          //Ending Arrow
          Leader leader = new Leader();
          leader.AppendVertex(endPoint);
          leader.AppendVertex(startPoint);
          leader.HasArrowHead = true;
          leader.Dimasz = 0.11;

          curSpace.AppendEntity(leader);
          tr.AddNewlyCreatedDBObject(leader, true);

          arrowPosition = new Point3d(arrowPosition.X + .2, arrowPosition.Y, endPoint.Z);
        }


        // Create a rectangle with a dotted line
        Point3d rectStart = new Point3d(InteriorPosition.X + .7, InteriorPosition.Y+.09, 0);
        Point3d rectEnd = new Point3d(arrowPosition.X, rectStart.Y - .57, 0);
        Autodesk.AutoCAD.DatabaseServices.Polyline rectangle = new Autodesk.AutoCAD.DatabaseServices.Polyline();
        rectangle.AddVertexAt(0, new Point2d(rectStart.X, rectStart.Y), 0, 0, 0);
        rectangle.AddVertexAt(1, new Point2d(rectEnd.X, rectStart.Y), 0, 0, 0);
        rectangle.AddVertexAt(2, new Point2d(rectEnd.X, rectEnd.Y), 0, 0, 0);
        rectangle.AddVertexAt(3, new Point2d(rectStart.X, rectEnd.Y), 0, 0, 0);
        rectangle.Closed = true;

        // Set the linetype to dotted
        LinetypeTable linetypeTable = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;
        if (linetypeTable.Has("DASHED")) {
          rectangle.Linetype = "DASHED";
        }
        else {
          ed.WriteMessage("\nLinetype 'DASHED' not found. Using continuous line.");
        }

        curSpace.AppendEntity(rectangle);
        tr.AddNewlyCreatedDBObject(rectangle, true);

        tr.Commit();
      }
    }


  }
}
