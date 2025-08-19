using System;
using System.Collections.Generic;
using System.Diagnostics.Metrics;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Accord.Statistics.Distributions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using ElectricalCommands.ElectricalEntity;
using GMEPElectricalCommands.GmepDatabase;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Editor = Autodesk.AutoCAD.EditorInput.Editor;

namespace ElectricalCommands.Lighting
{
  class LightingControlDiagram
  {
    LightingTimeClock TimeClock;
    Point3d InteriorPosition;
    Point3d ExteriorPosition;
    List<ElectricalEntity.LightingLocation> Locations;
    List<ElectricalEntity.LightingFixture> Fixtures;
    //List<ElectricalEntity.Equipment> Equipments;
    double SectionSeparation;

    public LightingControlDiagram(
      LightingTimeClock timeClock,
      List<ElectricalEntity.LightingLocation> locations,
      List<ElectricalEntity.LightingFixture> fixtures
       //List<ElectricalEntity.Equipment> equipments
    )
    {
      this.TimeClock = timeClock;
      this.Fixtures = fixtures;
      this.Locations = locations;
      //this.Equipments = equipments;
      this.SectionSeparation = 0.4;
      this.CreateDiagram();
    }

    public void CreateDiagram()
    {
      InitializeDiagram();
      GraphInteriorLighting();
      GraphExteriorLighting();
    }

    public void InitializeDiagram()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      GmepDatabase gmepDb = new GmepDatabase();
      string projectId = gmepDb.GetProjectId(CADObjectCommands.GetProjectNoFromFileName());
      Console.WriteLine("projectId" + projectId);
      List<ElectricalEntity.Panel> panels = gmepDb.GetPanels(projectId);
      ElectricalEntity.Panel panel = panels.Find(p => p.Id == TimeClock.AdjacentPanelId);
      string panelName = panel.Name;

      Point3d point;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord baseBlock;
        BlockReference br = CADObjectCommands.CreateBlockReference(tr, bt, "LTG CTRL BASE", out baseBlock, out point);
        Console.WriteLine($"X: {point.X}, Y: {point.Y}, Z: {point.Z}");
        if (br != null)
        {
          BlockTableRecord curSpace = (BlockTableRecord)
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
          curSpace.AppendEntity(br);
          tr.AddNewlyCreatedDBObject(br, true);

          foreach (ObjectId objId in baseBlock)
          {
            DBObject obj = tr.GetObject(objId, OpenMode.ForRead);
            AttributeDefinition attDef = obj as AttributeDefinition;
            if (attDef != null && !attDef.Constant)
            {
              using (AttributeReference attRef = new AttributeReference())
              {
                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                attRef.Position = attDef.Position.TransformBy(br.BlockTransform);
                if (attDef.Tag == "VOLTAGE")
                {
                  attRef.TextString = TimeClock.Voltage + "V";
                }
                if (attDef.Tag == "SWITCH")
                {
                  attRef.TextString = TimeClock.BypassSwitchName;
                }
                if (attDef.Tag == "LOCATION")
                {
                  attRef.TextString = TimeClock.BypassSwitchLocation.ToUpper();
                }
                if (attDef.Tag == "DESCRIPTION")
                {
                  attRef.TextString =
                    "("
                    + TimeClock.Voltage
                    + "V) ASTRONOMICAL 365-DAY PROGRAMMABLE TIME SWITCH \""
                    + TimeClock.Name
                    + "\" LOCATED ADJACENT TO PANEL \""
                    + panelName
                    + "\" IN A NEMA 1 ENCLOSURE. (TORK #ELC74 OR APPROVED EQUAL)\r\n";
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
          foreach (
            DynamicBlockReferenceProperty property in br.DynamicBlockReferencePropertyCollection
          )
          {
            if (property.PropertyName == "Position1 X")
            {
              ed.WriteMessage(property.Value.ToString());
              x = (double)property.Value;
            }
            if (property.PropertyName == "Position1 Y")
            {
              ed.WriteMessage(property.Value.ToString());
              y = (double)property.Value;
            }
            if (property.PropertyName == "Position2 X")
            {
              ed.WriteMessage(property.Value.ToString());
              x2 = (double)property.Value;
            }
            if (property.PropertyName == "Position2 Y")
            {
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

    public void GraphInteriorLighting()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      List<LightingLocation> indoorLocations = Locations
        .Where(location => !location.Outdoor)
        .ToList();

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord curSpace = (BlockTableRecord)
          tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
        bool isFirstFlag = true;
        foreach (LightingLocation location in indoorLocations)
        {
          if (!isFirstFlag)
          {
            try
            {
              Point3d startPoint = InteriorPosition;
              Point3d endPoint = new Point3d(startPoint.X, startPoint.Y - 1.75, startPoint.Z);
              Line verticalLine = new Line(startPoint, endPoint);
              verticalLine.Layer = "E-CND1";
              curSpace.AppendEntity(verticalLine);
              tr.AddNewlyCreatedDBObject(verticalLine, true);

              Circle circle = new Circle(endPoint, Vector3d.ZAxis, .020);
              circle.Layer = "E-CND1";
              curSpace.AppendEntity(circle);
              tr.AddNewlyCreatedDBObject(circle, true);

              Hatch hatch = new Hatch();
              hatch.Layer = "E-CND1";
              curSpace.AppendEntity(hatch);
              tr.AddNewlyCreatedDBObject(hatch, true);

              hatch.SetDatabaseDefaults();
              hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
              hatch.Associative = true;
              hatch.AppendLoop(HatchLoopTypes.Default, new ObjectIdCollection { circle.ObjectId });
              hatch.EvaluateHatch(true);

              InteriorPosition = endPoint;
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
              ed.WriteMessage($"\nError: {ex.Message}");
            }
          }
          isFirstFlag = false;

          GraphInteriorLocationSection(location);
        }

        tr.Commit();
      }
    }

    public void GraphExteriorLighting()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      List<LightingLocation> outdoorLocations = Locations
        .Where(location => location.Outdoor)
        .ToList();
      Console.WriteLine("outdoorLocations: "+ outdoorLocations.Count);
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord curSpace = (BlockTableRecord)
          tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

        //Setting Starting Circle and Starting point based on section separator
        Point3d NewExteriorPosition = new Point3d(
          ExteriorPosition.X + SectionSeparation,
          ExteriorPosition.Y,
          ExteriorPosition.Z
        );
        Line line = new Line(ExteriorPosition, NewExteriorPosition);
        line.Layer = "E-CND1";
        curSpace.AppendEntity(line);
        tr.AddNewlyCreatedDBObject(line, true);

        Circle circle1 = new Circle(NewExteriorPosition, Vector3d.ZAxis, .020);
        circle1.Layer = "E-CND1";
        curSpace.AppendEntity(circle1);
        tr.AddNewlyCreatedDBObject(circle1, true);

        Hatch hatch1 = new Hatch();
        hatch1.Layer = "E-CND1";
        curSpace.AppendEntity(hatch1);
        tr.AddNewlyCreatedDBObject(hatch1, true);
        hatch1.SetDatabaseDefaults();
        hatch1.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
        hatch1.Associative = true;
        hatch1.AppendLoop(HatchLoopTypes.Default, new ObjectIdCollection { circle1.ObjectId });
        hatch1.EvaluateHatch(true);

        ExteriorPosition = NewExteriorPosition;

        //Graphing Exterior Locations

        bool isFirstFlag = true;
        foreach (LightingLocation location in outdoorLocations)
        {
          Console.WriteLine("location: "+ location.LocationName);
          if (!isFirstFlag)
          {
            try
            {
              Point3d startPoint = ExteriorPosition;
              Point3d endPoint = new Point3d(startPoint.X, startPoint.Y - 1.75, startPoint.Z);
              Line verticalLine = new Line(startPoint, endPoint);
              verticalLine.Layer = "E-CND1";
              curSpace.AppendEntity(verticalLine);
              tr.AddNewlyCreatedDBObject(verticalLine, true);

              Circle circle = new Circle(endPoint, Vector3d.ZAxis, .020);
              circle.Layer = "E-CND1";
              curSpace.AppendEntity(circle);
              tr.AddNewlyCreatedDBObject(circle, true);

              Hatch hatch = new Hatch();
              hatch.Layer = "E-CND1";
              curSpace.AppendEntity(hatch);
              tr.AddNewlyCreatedDBObject(hatch, true);

              hatch.SetDatabaseDefaults();
              hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
              hatch.Associative = true;
              hatch.AppendLoop(HatchLoopTypes.Default, new ObjectIdCollection { circle.ObjectId });
              hatch.EvaluateHatch(true);

              ExteriorPosition = endPoint;
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
              ed.WriteMessage($"\nError: {ex.Message}");
            }
          }
          isFirstFlag = false;
          Console.WriteLine("location2: " + location.LocationName);
          GraphExteriorLocationSection(location);
        }

        tr.Commit();
      }
    }

    private void GraphInteriorLocationSection(LightingLocation location)
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      List<LightingFixture> fixturesAtLocation = Fixtures
        .Where(fixture => fixture.LocationId == location.Id)
        .ToList();
      Console.WriteLine("interiorFixtures: " + Fixtures.Count);

      List<LightingFixture> uniqueFixtures = new List<LightingFixture>();
      var fixtureDict = new Dictionary<(string ParentName, int Circuit), LightingFixture>();
      Console.WriteLine("fixturesAtLocation: ", fixturesAtLocation.Count);
      foreach (var fixture in fixturesAtLocation)
      {
        var combination = (fixture.ParentName, fixture.Circuit);
        if (fixtureDict.ContainsKey(combination))
        {
          // If a duplicate is found and it has EmCapable = true, replace the existing entry
          if (fixture.EmCapable)
          {
            fixtureDict[combination] = fixture;
          }
        }
        else
        {
          fixtureDict[combination] = fixture;
        }
      }

      // Convert the dictionary values to a list
      uniqueFixtures = fixtureDict.Values.ToList();

      //This method will graph the section for each interior lighting location
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord curSpace = (BlockTableRecord)
          tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
        TextStyleTable textStyleTable =
          tr.GetObject(db.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
        ObjectId textStyleId = textStyleTable["gmep"];

        //Graphing Conduit Loop
        Point3d startPoint = InteriorPosition;
        Point3d endPoint = new Point3d(startPoint.X + 1, startPoint.Y, startPoint.Z);
        Line horizontalLine = new Line(startPoint, endPoint);
        horizontalLine.Layer = "E-CND1";
        curSpace.AppendEntity(horizontalLine);
        tr.AddNewlyCreatedDBObject(horizontalLine, true);

        startPoint = endPoint;
        endPoint = new Point3d(endPoint.X, endPoint.Y - .1, endPoint.Z);
        Line verticalLine = new Line(startPoint, endPoint);
        verticalLine.Layer = "E-CND1";
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

        startPoint = new Point3d(endPoint.X, endPoint.Y - radius * 2, endPoint.Z);
        endPoint = new Point3d(startPoint.X, startPoint.Y - .1, startPoint.Z);
        Line verticalLine2 = new Line(startPoint, endPoint);
        verticalLine2.Layer = "E-CND1";
        curSpace.AppendEntity(verticalLine2);
        tr.AddNewlyCreatedDBObject(verticalLine2, true);

        startPoint = endPoint;
        endPoint = new Point3d(endPoint.X - .4, endPoint.Y, endPoint.Z);
        Line horizontalLine2 = new Line(startPoint, endPoint);
        horizontalLine2.Layer = "E-CND1";
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
        Point3d arrowPosition = new Point3d(
          InteriorPosition.X + 1.3,
          InteriorPosition.Y + .9,
          endPoint.Z
        );
        Point3d? emStartPosition = null;

        double tempSeparator = 0;
        Console.WriteLine("interioruniqueFixtures: "+uniqueFixtures.Count);
        for (int i = 0; i < uniqueFixtures.Count; i++) {
          LightingFixture fixture = uniqueFixtures[i];
        }
        foreach (LightingFixture fixture in uniqueFixtures)
        {
          double offsetX = 0.2;
          Console.WriteLine($"{fixture.ParentName}-{fixture.Circuit} (EM: {fixture.EmCapable})");
          //Begin Arrow
          startPoint = arrowPosition;
          endPoint = new Point3d(startPoint.X, startPoint.Y - 1.06, startPoint.Z);
          Line beginArrow = new Line(startPoint, endPoint);
          beginArrow.Layer = "E-CND1";
          curSpace.AppendEntity(beginArrow);
          tr.AddNewlyCreatedDBObject(beginArrow, true);

          
          Point3d v2Start = new Point3d(startPoint.X + offsetX, startPoint.Y, startPoint.Z);
          Point3d v2End = new Point3d(startPoint.X + offsetX, startPoint.Y - 1.06, startPoint.Z);
          Line line2 = new Line(v2Start, v2End);
          line2.Layer = "E-CND1";
          curSpace.AppendEntity(line2);
          tr.AddNewlyCreatedDBObject(line2, true);
          //EM Circle
          if (fixture.EmCapable)
          {
            Point3d emStartPoint = new Point3d(arrowPosition.X, arrowPosition.Y - .31, -1); // Start point for the EM section
            Circle EmCircle = new Circle(emStartPoint, Vector3d.ZAxis, .030);
            EmCircle.Layer = "E-CND1";
            curSpace.AppendEntity(EmCircle);
            tr.AddNewlyCreatedDBObject(EmCircle, true);
            Hatch hatch = new Hatch();
            hatch.Layer = "E-CND1";
            curSpace.AppendEntity(hatch);
            tr.AddNewlyCreatedDBObject(hatch, true);
            hatch.SetDatabaseDefaults();
            hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            hatch.Associative = true;
            hatch.AppendLoop(HatchLoopTypes.Default, new ObjectIdCollection { EmCircle.ObjectId });
            hatch.EvaluateHatch(true);
            if (emStartPosition == null)
            {
              emStartPosition = emStartPoint;
            }
          }
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
          //offset label
          DBText label2 = new DBText();
          label2.Position = new Point3d(startPoint.X, startPoint.Y, startPoint.Z);
          label2.Rotation = (Math.PI / 2);
          label2.Height = radius * .9;
          label2.TextString = fixture.ParentName + "-" + fixture.Circuit.ToString();
          label2.HorizontalMode = TextHorizontalMode.TextCenter;
          label2.VerticalMode = TextVerticalMode.TextVerticalMid;
          label2.AlignmentPoint = new Point3d(startPoint.X + offsetX, startPoint.Y, startPoint.Z);
          label2.Justify = AttachmentPoint.BottomRight;
          label2.Layer = "E-TEXT";
          curSpace.AppendEntity(label2);
          tr.AddNewlyCreatedDBObject(label2, true);
          //Draw Horizontal lines
          Line separator = new Line(
            new Point3d(endPoint.X - .07, endPoint.Y, endPoint.Z),
            new Point3d(endPoint.X + .07, endPoint.Y, endPoint.Z)
          );
          separator.Layer = "E-TEXT";
          curSpace.AppendEntity(separator);
          tr.AddNewlyCreatedDBObject(separator, true);

          //offset s1
          Line separator1Off = new Line(
            new Point3d(endPoint.X - .07 + offsetX, endPoint.Y, endPoint.Z),
            new Point3d(endPoint.X + .07 + offsetX, endPoint.Y, endPoint.Z)
          );
          separator1Off.Layer = "E-TEXT";
          curSpace.AppendEntity(separator1Off);
          tr.AddNewlyCreatedDBObject(separator1Off, true);

          //s2
          startPoint = new Point3d(endPoint.X, endPoint.Y - .05, endPoint.Z);
          endPoint = new Point3d(startPoint.X, startPoint.Y - .4, startPoint.Z);

          Line separator2 = new Line(
            new Point3d(startPoint.X - .07, startPoint.Y, startPoint.Z),
            new Point3d(startPoint.X + .07, startPoint.Y, startPoint.Z)
          );
          separator2.Layer = "E-TEXT";
          curSpace.AppendEntity(separator2);
          tr.AddNewlyCreatedDBObject(separator2, true);
          //offset s2
          Line separator2Off = new Line(
            new Point3d(startPoint.X - .07 + offsetX, startPoint.Y, startPoint.Z),
            new Point3d(startPoint.X + .07 + offsetX, startPoint.Y, startPoint.Z)
          );
          separator2Off.Layer = "E-TEXT";
          curSpace.AppendEntity(separator2Off);
          tr.AddNewlyCreatedDBObject(separator2Off, true);
          //Ending Arrow
          Leader leader = new Leader();
          leader.Layer = "E-CND1";
          leader.AppendVertex(endPoint);
          leader.AppendVertex(startPoint);
          leader.HasArrowHead = true;
          leader.Dimasz = 0.11;

          curSpace.AppendEntity(leader);
          tr.AddNewlyCreatedDBObject(leader, true);

          arrowPosition = new Point3d(arrowPosition.X + .2, arrowPosition.Y, endPoint.Z);
          //offset ending arrow
          Leader leader2 = new Leader();
          leader2.Layer = "E-CND1";
          leader2.HasArrowHead = true;
          leader2.AppendVertex(new Point3d(endPoint.X + offsetX, endPoint.Y, endPoint.Z));
          leader2.AppendVertex(new Point3d(startPoint.X + offsetX, startPoint.Y, startPoint.Z));
          curSpace.AppendEntity(leader2);
          tr.AddNewlyCreatedDBObject(leader2, true);
          leader2.Dimasz = 0.11;
          //Adjust Separator
          tempSeparator += .2;
        }
        // Draw the final EM leader line if applicable
        if (emStartPosition != null)
        {
          Point3d emStartPosition2 = (Point3d)emStartPosition;
          Point3d leaderEndPoint = new Point3d(arrowPosition.X + .3, emStartPosition2.Y, 0);
          Leader emLeader = new Leader();
          emLeader.AppendVertex(leaderEndPoint);
          emLeader.AppendVertex(emStartPosition2);
          emLeader.HasArrowHead = true;
          emLeader.Dimasz = 0.11;
          emLeader.Layer = "E-TEXT";
          curSpace.AppendEntity(emLeader);
          tr.AddNewlyCreatedDBObject(emLeader, true);

          DBText EmLabel = new DBText();
          EmLabel.Position = new Point3d(
            leaderEndPoint.X + .05,
            leaderEndPoint.Y,
            leaderEndPoint.Z
          );
          EmLabel.Height = radius * .9;
          EmLabel.TextString = "TO EM LIGHT";
          EmLabel.HorizontalMode = TextHorizontalMode.TextCenter;
          EmLabel.VerticalMode = TextVerticalMode.TextVerticalMid;
          EmLabel.AlignmentPoint = new Point3d(
            leaderEndPoint.X + .05,
            leaderEndPoint.Y,
            leaderEndPoint.Z
          );
          EmLabel.Justify = AttachmentPoint.TopLeft;
          emLeader.Dimasz = 0.11;
          EmLabel.Layer = "E-TEXT";
          EmLabel.TextStyleId = textStyleId;
          curSpace.AppendEntity(EmLabel);
          tr.AddNewlyCreatedDBObject(EmLabel, true);
          //Adjust Separator
          tempSeparator += 1;
        }
        if (tempSeparator > SectionSeparation)
        {
          SectionSeparation = tempSeparator;
        }
        // Create a rectangle with a dotted line
        double offset = 0.2;
        Point3d rectStart = new Point3d(InteriorPosition.X + .7, InteriorPosition.Y + .09, 0);
        Point3d rectEnd = new Point3d(arrowPosition.X + offset ,rectStart.Y - .57, 0);
        Autodesk.AutoCAD.DatabaseServices.Polyline rectangle =
          new Autodesk.AutoCAD.DatabaseServices.Polyline();
        rectangle.AddVertexAt(0, new Point2d(rectStart.X, rectStart.Y), 0, 0, 0);
        rectangle.AddVertexAt(1, new Point2d(rectEnd.X, rectStart.Y), 0, 0, 0);
        rectangle.AddVertexAt(2, new Point2d(rectEnd.X, rectEnd.Y), 0, 0, 0);
        rectangle.AddVertexAt(3, new Point2d(rectStart.X, rectEnd.Y), 0, 0, 0);
        rectangle.Layer = "E-TEXT";
        rectangle.Closed = true;
        // Set the linetype to dotted
        LinetypeTable linetypeTable =
          tr.GetObject(db.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;
        if (linetypeTable.Has("HIDDEN"))
        {
          rectangle.Linetype = "HIDDEN";
        }
        else
        {
          ed.WriteMessage("\nLinetype 'DASHED2' not found. Using continuous line.");
        }
        curSpace.AppendEntity(rectangle);
        tr.AddNewlyCreatedDBObject(rectangle, true);
        //Append Location Text
        DBText locationLabel = new DBText();
        locationLabel.Position = new Point3d(
          InteriorPosition.X + 1.2,
          InteriorPosition.Y - .65,
          InteriorPosition.Z
        );
        locationLabel.Height = radius * .9;
        locationLabel.TextString = location.LocationName;
        locationLabel.HorizontalMode = TextHorizontalMode.TextCenter;
        locationLabel.VerticalMode = TextVerticalMode.TextVerticalMid;
        locationLabel.AlignmentPoint = new Point3d(
          InteriorPosition.X + 1.2,
          InteriorPosition.Y - .65,
          InteriorPosition.Z
        );
        locationLabel.Justify = AttachmentPoint.TopLeft;
        locationLabel.Layer = "E-TEXT";
        locationLabel.TextStyleId = textStyleId;
        curSpace.AppendEntity(locationLabel);
        tr.AddNewlyCreatedDBObject(locationLabel, true);

        tr.Commit();
      }
    }

    private void GraphExteriorLocationSection(LightingLocation location)
    {
      Console.WriteLine("location3: " + location.LocationName + location.Id);
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      List<LightingFixture> fixturesAtLocation = Fixtures
        .Where(fixture => fixture.LocationId == location.Id)
        .ToList();
      Console.WriteLine($"Fixture: "+ Fixtures.Count);
      for (int i = 0; i < Fixtures.Count; i++) {
        Console.WriteLine($"Fixture {Fixtures[i].LocationId}");
      }
      Console.WriteLine($"fixturesAtLocation:" + fixturesAtLocation.Count);
      
      for (int i = 0; i < Fixtures.Count; i++) {
        Console.WriteLine($"Fixture {Fixtures[i].LocationId}");
      }
      List<LightingFixture> uniqueFixtures = new List<LightingFixture>();
      var fixtureDict = new Dictionary<(string ParentName, int Circuit), LightingFixture>();

      foreach (var fixture in fixturesAtLocation)
      {
        var combination = (fixture.ParentName, fixture.Circuit);
        if (fixtureDict.ContainsKey(combination))
        {
          // If a duplicate is found and it has EmCapable = true, replace the existing entry
          if (fixture.EmCapable)
          {
            fixtureDict[combination] = fixture;
          }
        }
        else
        {
          fixtureDict[combination] = fixture;
        }
      }
      Console.WriteLine($"Fixtures dict :"+ fixtureDict.Count);
      Console.WriteLine($"Fixtures at Location {location.Id}: {fixturesAtLocation.Count}");
        // Convert the dictionary values to a list
        uniqueFixtures = fixtureDict.Values.ToList();
      Console.WriteLine("uniqueFixtures Count: " + uniqueFixtures.Count);

      //This method will graph the section for each interior lighting location
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord curSpace = (BlockTableRecord)
          tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
        TextStyleTable textStyleTable =
          tr.GetObject(db.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
        ObjectId textStyleId = textStyleTable["Standard"];

        //Graphing Conduit Loop
        Point3d startPoint = ExteriorPosition;
        Point3d endPoint = new Point3d(startPoint.X + 1, startPoint.Y, startPoint.Z);
        Line horizontalLine = new Line(startPoint, endPoint);
        horizontalLine.Layer = "E-CND1";
        curSpace.AppendEntity(horizontalLine);
        tr.AddNewlyCreatedDBObject(horizontalLine, true);

        startPoint = endPoint;
        //endPoint = new Point3d(endPoint.X, endPoint.Y - .1, endPoint.Z);
        //Line verticalLine = new Line(startPoint, endPoint);
        //curSpace.AppendEntity(verticalLine);
        //tr.AddNewlyCreatedDBObject(verticalLine, true);

        double radius = .09;
        Point3d circleCenter = new Point3d(endPoint.X + radius, endPoint.Y, endPoint.Z);
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

        startPoint = new Point3d(endPoint.X + radius, endPoint.Y - radius, endPoint.Z);
        endPoint = new Point3d(startPoint.X, startPoint.Y - .1, startPoint.Z);
        Line verticalLine2 = new Line(startPoint, endPoint);
        verticalLine2.Layer = "E-CND1";
        curSpace.AppendEntity(verticalLine2);
        tr.AddNewlyCreatedDBObject(verticalLine2, true);

        startPoint = endPoint;
        endPoint = new Point3d(endPoint.X - .4, endPoint.Y, endPoint.Z);
        Line horizontalLine2 = new Line(startPoint, endPoint);
        horizontalLine2.Layer = "E-CND1";
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
        Point3d arrowPosition = new Point3d(
          ExteriorPosition.X + 1.3,
          ExteriorPosition.Y + .9,
          endPoint.Z
        );
        Point3d? emStartPosition = null;
        double tempSeparator = 0;
        double offsetX = 0.2;
        foreach (LightingFixture fixture in uniqueFixtures)
        {
          Console.WriteLine($"{fixture.ParentName}-{fixture.Circuit} (EM: {fixture.EmCapable})");
          //Begin Arrow
          startPoint = arrowPosition;
          endPoint = new Point3d(startPoint.X, startPoint.Y - 1, startPoint.Z);
          Console.WriteLine(startPoint.Y - 1);
          Line beginArrow = new Line(startPoint, endPoint);
          beginArrow.Layer = "E-CND1";
          curSpace.AppendEntity(beginArrow);
          tr.AddNewlyCreatedDBObject(beginArrow, true);

          
          Point3d v2Start = new Point3d(startPoint.X + offsetX, startPoint.Y, startPoint.Z);
          Point3d v2End = new Point3d(startPoint.X + offsetX, startPoint.Y - 1.06, startPoint.Z);
          Line line2 = new Line(v2Start, v2End);
          line2.Layer = "E-CND1";
          curSpace.AppendEntity(line2);
          tr.AddNewlyCreatedDBObject(line2, true);
          //EM Circle
          if (fixture.EmCapable)
          {
            Point3d emStartPoint = new Point3d(arrowPosition.X, arrowPosition.Y - .31, -1); // Start point for the EM section
            Circle EmCircle = new Circle(emStartPoint, Vector3d.ZAxis, .030);
            EmCircle.Layer = "E-CND1";
            curSpace.AppendEntity(EmCircle);
            tr.AddNewlyCreatedDBObject(EmCircle, true);
            Hatch hatch = new Hatch();
            hatch.Layer = "E-CND1";
            curSpace.AppendEntity(hatch);
            tr.AddNewlyCreatedDBObject(hatch, true);
            hatch.SetDatabaseDefaults();
            hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            hatch.Associative = true;
            hatch.AppendLoop(HatchLoopTypes.Default, new ObjectIdCollection { EmCircle.ObjectId });
            hatch.EvaluateHatch(true);
            if (emStartPosition == null)
            {
              emStartPosition = emStartPoint;
            }
          }
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
          //offset label
          DBText label2 = new DBText();
          label2.Position = new Point3d(startPoint.X, startPoint.Y, startPoint.Z);
          label2.Rotation = (Math.PI / 2);
          label2.Height = radius * .9;
          label2.TextString = fixture.ParentName + "-" + fixture.Circuit.ToString();
          label2.HorizontalMode = TextHorizontalMode.TextCenter;
          label2.VerticalMode = TextVerticalMode.TextVerticalMid;
          label2.AlignmentPoint = new Point3d(startPoint.X + offsetX, startPoint.Y, startPoint.Z);
          label2.Justify = AttachmentPoint.BottomRight;
          label2.Layer = "E-TEXT";
          curSpace.AppendEntity(label2);
          tr.AddNewlyCreatedDBObject(label2, true);
          //Draw Horizontal lines
          Line separator = new Line(
            new Point3d(endPoint.X - .07, endPoint.Y, endPoint.Z),
            new Point3d(endPoint.X + .07, endPoint.Y, endPoint.Z)
          );
          separator.Layer = "E-TEXT";
          curSpace.AppendEntity(separator);
          tr.AddNewlyCreatedDBObject(separator, true);
          //offset s1
          Line separator1Off = new Line(
            new Point3d(endPoint.X - .07 + offsetX, endPoint.Y, endPoint.Z),
            new Point3d(endPoint.X + .07 + offsetX, endPoint.Y, endPoint.Z)
          );
          separator1Off.Layer = "E-TEXT";
          curSpace.AppendEntity(separator1Off);
          tr.AddNewlyCreatedDBObject(separator1Off, true);

          startPoint = new Point3d(endPoint.X, endPoint.Y - .05, endPoint.Z);
          endPoint = new Point3d(startPoint.X, startPoint.Y - .4, startPoint.Z);

          Line separator2 = new Line(
            new Point3d(startPoint.X - .07, startPoint.Y, startPoint.Z),
            new Point3d(startPoint.X + .07, startPoint.Y, startPoint.Z)
          );
          separator2.Layer = "E-TEXT";
          curSpace.AppendEntity(separator2);
          tr.AddNewlyCreatedDBObject(separator2, true);
          //offset s2
          Line separator2Off = new Line(
            new Point3d(startPoint.X - .07 + offsetX, startPoint.Y, startPoint.Z),
            new Point3d(startPoint.X + .07 + offsetX, startPoint.Y, startPoint.Z)
          );
          separator2Off.Layer = "E-TEXT";
          curSpace.AppendEntity(separator2Off);
          tr.AddNewlyCreatedDBObject(separator2Off, true);
          //Ending Arrow
          Leader leader = new Leader();
          leader.Layer = "E-CND1";
          leader.AppendVertex(endPoint);
          leader.AppendVertex(startPoint);
          leader.HasArrowHead = true;
          leader.Dimasz = 0.11;

          curSpace.AppendEntity(leader);
          tr.AddNewlyCreatedDBObject(leader, true);

          arrowPosition = new Point3d(arrowPosition.X + .2, arrowPosition.Y, endPoint.Z);
        }
        //offset ending arrow
        Leader leader2 = new Leader();
        leader2.Layer = "E-CND1";
        leader2.HasArrowHead = true;
        leader2.AppendVertex(new Point3d(endPoint.X + offsetX, endPoint.Y, endPoint.Z));
        leader2.AppendVertex(new Point3d(startPoint.X + offsetX, startPoint.Y, startPoint.Z));
        curSpace.AppendEntity(leader2);
        tr.AddNewlyCreatedDBObject(leader2, true);
        leader2.Dimasz = 0.11;
        //Adjust Separator
        tempSeparator += .2;
        // Draw the final EM leader line if applicable
        if (emStartPosition != null)
        {
          Point3d emStartPosition2 = (Point3d)emStartPosition;
          Point3d leaderEndPoint = new Point3d(arrowPosition.X + .3, emStartPosition2.Y, 0);
          Leader emLeader = new Leader();
          emLeader.AppendVertex(leaderEndPoint);
          emLeader.AppendVertex(emStartPosition2);
          emLeader.HasArrowHead = true;
          emLeader.Dimasz = 0.11;
          emLeader.Layer = "E-TEXT";
          curSpace.AppendEntity(emLeader);
          tr.AddNewlyCreatedDBObject(emLeader, true);

          DBText EmLabel = new DBText();
          EmLabel.Position = new Point3d(
            leaderEndPoint.X + .05,
            leaderEndPoint.Y,
            leaderEndPoint.Z
          );
          EmLabel.Height = radius * .9;
          EmLabel.TextString = "TO EM LIGHT";
          EmLabel.HorizontalMode = TextHorizontalMode.TextCenter;
          EmLabel.VerticalMode = TextVerticalMode.TextVerticalMid;
          EmLabel.AlignmentPoint = new Point3d(
            leaderEndPoint.X + .05,
            leaderEndPoint.Y,
            leaderEndPoint.Z
          );
          EmLabel.Justify = AttachmentPoint.TopLeft;
          EmLabel.Layer = "E-TEXT";
          curSpace.AppendEntity(EmLabel);
          tr.AddNewlyCreatedDBObject(EmLabel, true);
        }
        //Adjust Separator
        tempSeparator += 1;
        if (tempSeparator > SectionSeparation) {
          SectionSeparation = tempSeparator;
        }
        // Create a rectangle with a dotted line
        double offset = 0.2;
        Point3d rectStart = new Point3d(ExteriorPosition.X + .8, ExteriorPosition.Y + .15, 0);
        Point3d rectEnd = new Point3d(arrowPosition.X + offset, rectStart.Y - .50, 0);
        Autodesk.AutoCAD.DatabaseServices.Polyline rectangle =
          new Autodesk.AutoCAD.DatabaseServices.Polyline();
        rectangle.AddVertexAt(0, new Point2d(rectStart.X, rectStart.Y), 0, 0, 0);
        rectangle.AddVertexAt(1, new Point2d(rectEnd.X, rectStart.Y), 0, 0, 0);
        rectangle.AddVertexAt(2, new Point2d(rectEnd.X, rectEnd.Y), 0, 0, 0);
        rectangle.AddVertexAt(3, new Point2d(rectStart.X, rectEnd.Y), 0, 0, 0);
        rectangle.Layer = "E-TEXT";
        rectangle.Closed = true;
        // Set the linetype to dotted
        LinetypeTable linetypeTable =
          tr.GetObject(db.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;
        if (linetypeTable.Has("DASHED2"))
        {
          rectangle.Linetype = "DASHED2";
        }
        else
        {
          ed.WriteMessage("\nLinetype 'DASHED2' not found. Using continuous line.");
        }
        curSpace.AppendEntity(rectangle);
        tr.AddNewlyCreatedDBObject(rectangle, true);
        //Append Location Text
        DBText locationLabel = new DBText();
        locationLabel.Position = new Point3d(
          ExteriorPosition.X + 1.2,
          ExteriorPosition.Y - .60,
          ExteriorPosition.Z
        );
        locationLabel.Height = radius * .9;
        locationLabel.TextString = location.LocationName;
        locationLabel.HorizontalMode = TextHorizontalMode.TextCenter;
        locationLabel.VerticalMode = TextVerticalMode.TextVerticalMid;
        locationLabel.AlignmentPoint = new Point3d(
          ExteriorPosition.X + 1.2,
          ExteriorPosition.Y - .60,
          ExteriorPosition.Z
        );
        locationLabel.Justify = AttachmentPoint.TopLeft;
        locationLabel.Layer = "E-TEXT";
        curSpace.AppendEntity(locationLabel);
        tr.AddNewlyCreatedDBObject(locationLabel, true);

        tr.Commit();
      }
    }
  }
}
