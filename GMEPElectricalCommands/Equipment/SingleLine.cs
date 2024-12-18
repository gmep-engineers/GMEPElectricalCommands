using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands.Equipment
{
  public class SingleLine
  {
    public double width;
    public string type;
    public string name;
    public string id;
    public List<SingleLine> children;
    public Point3d startingPoint;
    public Point3d endingPoint;

    public SingleLine()
    {
      children = new List<SingleLine>();
    }

    public double AggregateWidths()
    {
      double sum = width;
      foreach (SingleLine child in children)
      {
        sum += child.AggregateWidths();
      }
      width = sum;
      return sum;
    }

    public virtual void SetChildStartingPoints(Point3d startingPoint)
    {
      this.startingPoint = startingPoint;

      foreach (SingleLine child in children)
      {
        child.SetChildStartingPoints(startingPoint);
      }
    }

    public void SetChildEndingPoint(Point3d endingPoint)
    {
      this.endingPoint = endingPoint;
    }

    public virtual void Make()
    {
      foreach (var child in children)
      {
        child.Make();
      }
    }
  }

  public class SLServiceFeeder : SingleLine
  {
    public bool isMultiMeter;
    public int amp;

    public SLServiceFeeder(string id, string name, bool isMultiMeter, int amp)
    {
      type = "service feeder";
      width = 2.5;
      this.name = name;
      this.id = id;
      this.isMultiMeter = isMultiMeter;
      this.amp = amp;
    }

    public override void SetChildStartingPoints(Point3d startingPoint)
    {
      double offset = 0;
      this.startingPoint = startingPoint;
      foreach (SingleLine child in children)
      {
        child.SetChildStartingPoints(
          new Point3d(startingPoint.X + 2.5 + offset, startingPoint.Y, startingPoint.Z)
        );
        offset += width + 1;
      }
    }

    public override void Make()
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        LineData boxLine1 = new LineData();
        boxLine1.Layer = "E-SYM1";
        boxLine1.StartPoint = new SimpleVector3d();
        boxLine1.EndPoint = new SimpleVector3d();
        boxLine1.StartPoint.X = startingPoint.X;
        boxLine1.StartPoint.Y = startingPoint.Y;
        boxLine1.EndPoint.X = startingPoint.X + 2.5;
        boxLine1.EndPoint.Y = startingPoint.Y;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, boxLine1, 1, "HIDDEN");

        LineData boxLine2 = new LineData();
        boxLine2.Layer = "E-SYM1";
        boxLine2.StartPoint = new SimpleVector3d();
        boxLine2.EndPoint = new SimpleVector3d();
        boxLine2.StartPoint.X = startingPoint.X;
        boxLine2.StartPoint.Y = startingPoint.Y;
        boxLine2.EndPoint.X = startingPoint.X;
        boxLine2.EndPoint.Y = startingPoint.Y - 2;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, boxLine2, 1, "HIDDEN");

        LineData boxLine3 = new LineData();
        boxLine3.Layer = "E-SYM1";
        boxLine3.StartPoint = new SimpleVector3d();
        boxLine3.EndPoint = new SimpleVector3d();
        boxLine3.StartPoint.X = startingPoint.X;
        boxLine3.StartPoint.Y = startingPoint.Y - 2;
        boxLine3.EndPoint.X = startingPoint.X + 2.5;
        boxLine3.EndPoint.Y = startingPoint.Y - 2;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, boxLine3, 1, "HIDDEN");

        LineData boxLine4 = new LineData();
        boxLine4.Layer = "E-SYM1";
        boxLine4.StartPoint = new SimpleVector3d();
        boxLine4.EndPoint = new SimpleVector3d();
        boxLine4.StartPoint.X = startingPoint.X + 0.75;
        boxLine4.StartPoint.Y = startingPoint.Y;
        boxLine4.EndPoint.X = startingPoint.X + 0.75;
        boxLine4.EndPoint.Y = startingPoint.Y - 2;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, boxLine4, 1, "HIDDEN");

        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = startingPoint.X + 0.375;
        conduitLine1.StartPoint.Y = startingPoint.Y - 0.2188;
        conduitLine1.EndPoint.X = startingPoint.X + 0.375;
        conduitLine1.EndPoint.Y = startingPoint.Y - 2;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);

        LineData conduitLine2 = new LineData();
        conduitLine2.Layer = "E-CND1";
        conduitLine2.StartPoint = new SimpleVector3d();
        conduitLine2.EndPoint = new SimpleVector3d();
        conduitLine2.StartPoint.X = startingPoint.X + 0.375;
        conduitLine2.StartPoint.Y = startingPoint.Y - 0.2188;
        conduitLine2.EndPoint.X = startingPoint.X + 1.25;
        conduitLine2.EndPoint.Y = startingPoint.Y - 0.2188;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);

        LineData conduitLine3 = new LineData();
        conduitLine3.Layer = "E-CND1";
        conduitLine3.StartPoint = new SimpleVector3d();
        conduitLine3.EndPoint = new SimpleVector3d();
        conduitLine3.StartPoint.X = startingPoint.X + 1.25;
        conduitLine3.StartPoint.Y = startingPoint.Y - 1.5;
        conduitLine3.EndPoint.X = startingPoint.X + 2;
        conduitLine3.EndPoint.Y = startingPoint.Y - 1.5;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine3, 1);

        LineData conduitLine4 = new LineData();
        conduitLine4.Layer = "E-CND1";
        conduitLine4.StartPoint = new SimpleVector3d();
        conduitLine4.EndPoint = new SimpleVector3d();
        conduitLine4.StartPoint.X = startingPoint.X + 2;
        conduitLine4.StartPoint.Y = startingPoint.Y - 1.5;
        conduitLine4.EndPoint.X = startingPoint.X + 2;
        conduitLine4.EndPoint.Y = startingPoint.Y - 0.2188;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine4, 1);

        LineData conduitLine5 = new LineData();
        conduitLine5.Layer = "E-CND1";
        conduitLine5.StartPoint = new SimpleVector3d();
        conduitLine5.EndPoint = new SimpleVector3d();
        conduitLine5.StartPoint.X = startingPoint.X + 2;
        conduitLine5.StartPoint.Y = startingPoint.Y - 0.2188;
        conduitLine5.EndPoint.X = startingPoint.X + 2.25;
        conduitLine5.EndPoint.Y = startingPoint.Y - 0.2188;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine5, 1);

        LineData feederLine = new LineData();
        feederLine.Layer = "E-CND1";
        feederLine.StartPoint = new SimpleVector3d();
        feederLine.EndPoint = new SimpleVector3d();
        feederLine.StartPoint.X = startingPoint.X + 0.375;
        feederLine.StartPoint.Y = startingPoint.Y - 2;
        feederLine.EndPoint.X = startingPoint.X + 0.375;
        feederLine.EndPoint.Y = startingPoint.Y - 2.875;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, feederLine, 1, "HIDDEN2");
        ObjectId arrowSymbol = bt["DOWN ARROW (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(startingPoint.X + 0.375, startingPoint.Y - 2.875, 0),
            arrowSymbol
          )
        )
        {
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        ObjectId dash1 = bt["SERVICE FEEDER DASH (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(startingPoint.X + 0.375, startingPoint.Y - 0.5, 0),
            dash1
          )
        )
        {
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        ObjectId dash2 = bt["SERVICE FEEDER DASH (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(startingPoint.X + 0.375, startingPoint.Y - 1.5, 0),
            dash2
          )
        )
        {
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }

        ObjectId gndBus = bt["GND BUS (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(startingPoint.X + 1.125, startingPoint.Y - 1.875, 0),
            gndBus
          )
        )
        {
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }

        tr.Commit();
      }
      foreach (var child in children)
      {
        child.Make();
      }
    }
  }

  public class SLMainBreakerSection : SingleLine
  {
    public int breakerSize;

    public SLMainBreakerSection(string id, string name)
    {
      type = "main breaker section";
      width = 1;
      this.name = name;
      this.id = id;
    }
  }

  public class SLDistributionSection : SingleLine
  {
    public SLDistributionSection(string id, string name)
    {
      type = "distribution section";
      width = 1;
      this.name = name;
      this.id = id;
    }
  }

  public class SLPanel : SingleLine
  {
    public bool isDistribution;
    public bool hasMeter;
    public bool hasCts;
    public bool hasGfp;
    public int distributionBreakerSize;
    public int mainBreakerSize;
    public int parentDistance;
    public string conduitSize;
    public string wireSize;
    public string voltageDrop;

    public SLPanel(string id, string name, bool isDistribution, bool hasMeter, int parentDistance)
    {
      type = "panel";
      width = 2;
      this.name = name;
      this.id = id;
      this.isDistribution = isDistribution;
      this.hasMeter = hasMeter;
      this.parentDistance = parentDistance;
    }

    public override void SetChildStartingPoints(Point3d startingPoint)
    {
      this.startingPoint = startingPoint;
      if (isDistribution)
      {
        double offset = 0;
        foreach (SingleLine child in children)
        {
          child.SetChildEndingPoint(
            new Point3d(
              startingPoint.X + (child.width / 2) + offset,
              startingPoint.Y - 5,
              startingPoint.Z
            )
          );
          child.SetChildStartingPoints(
            new Point3d(
              startingPoint.X + (child.width / 2) + offset,
              startingPoint.Y - 0.25,
              startingPoint.Z
            )
          );
          offset += child.width;
        }
      }
      else
      {
        int index = 0;
        foreach (SingleLine child in children)
        {
          if (index == 0)
          {
            child.SetChildEndingPoint(
              new Point3d(endingPoint.X + (child.width / 2), endingPoint.Y - 3.25, 0)
            );
            child.SetChildStartingPoints(
              new Point3d(endingPoint.X + (5 / 16), endingPoint.Y - (7 / 8), 0)
            );
          }
          if (index == 1)
          {
            child.SetChildEndingPoint(
              new Point3d(endingPoint.X - (child.width / 2), endingPoint.Y - 3.25, 0)
            );
            child.SetChildStartingPoints(
              new Point3d(endingPoint.X - (5 / 16), endingPoint.Y - (7 / 8), 0)
            );
          }
          if (index == 2)
          {
            child.SetChildEndingPoint(
              new Point3d(endingPoint.X + child.width, endingPoint.Y - 3.25, 0)
            );
            child.SetChildStartingPoints(
              new Point3d(endingPoint.X + (5 / 16), endingPoint.Y - (1 / 8), 0)
            );
          }
          if (index == 3)
          {
            child.SetChildEndingPoint(
              new Point3d(endingPoint.X - child.width, endingPoint.Y - 3.25, 0)
            );
            child.SetChildStartingPoints(
              new Point3d(endingPoint.X - (5 / 16), endingPoint.Y - (1 / 8), 0)
            );
          }
        }
      }
    }

    public override void Make()
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        if (isDistribution)
        {
          LineData lineData1 = new LineData();
          lineData1.Layer = "E-SYM1";
          lineData1.StartPoint = new SimpleVector3d();
          lineData1.EndPoint = new SimpleVector3d();
          lineData1.StartPoint.X = startingPoint.X;
          lineData1.StartPoint.Y = startingPoint.Y;
          lineData1.EndPoint.X = startingPoint.X + width - 2;
          lineData1.EndPoint.Y = startingPoint.Y;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, lineData1, 1, "HIDDEN");

          LineData lineData2 = new LineData();
          lineData2.Layer = "E-SYM1";
          lineData2.StartPoint = new SimpleVector3d();
          lineData2.EndPoint = new SimpleVector3d();
          lineData2.StartPoint.X = startingPoint.X;
          lineData2.StartPoint.Y = startingPoint.Y;
          lineData2.EndPoint.X = startingPoint.X;
          lineData2.EndPoint.Y = startingPoint.Y - 2;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, lineData2, 1, "HIDDEN");

          LineData lineData3 = new LineData();
          lineData3.Layer = "E-SYM1";
          lineData3.StartPoint = new SimpleVector3d();
          lineData3.EndPoint = new SimpleVector3d();
          lineData3.StartPoint.X = startingPoint.X;
          lineData3.StartPoint.Y = startingPoint.Y - 2;
          lineData3.EndPoint.X = startingPoint.X + width - 2;
          lineData3.EndPoint.Y = startingPoint.Y - 2;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, lineData3, 1, "HIDDEN");
          Polyline2dData polyData = new Polyline2dData();

          polyData.Layer = "E-CND1";
          polyData.Vertices.Add(
            new SimpleVector3d(startingPoint.X - 0.25, startingPoint.Y - 0.1875, 0)
          );
          polyData.Vertices.Add(
            new SimpleVector3d(startingPoint.X - 0.25, startingPoint.Y - 0.25, 0)
          );
          polyData.Vertices.Add(
            new SimpleVector3d(startingPoint.X + width - 2.25, startingPoint.Y - 0.25, 0)
          );
          polyData.Vertices.Add(
            new SimpleVector3d(startingPoint.X + width - 2.25, startingPoint.Y - 0.1875, 0)
          );
          polyData.Vertices.Add(
            new SimpleVector3d(startingPoint.X - 0.25, startingPoint.Y - 0.1875, 0)
          );
          polyData.Closed = true;
          CADObjectCommands.CreatePolyline2d(new Point3d(), tr, btr, polyData, 1);

          if (hasMeter)
          {
            if (hasCts)
            {
              LineData conduitLine1 = new LineData();
              conduitLine1.Layer = "E-CND1";
              conduitLine1.StartPoint = new SimpleVector3d();
              conduitLine1.EndPoint = new SimpleVector3d();
              conduitLine1.StartPoint.X = startingPoint.X - 1.25;
              conduitLine1.StartPoint.Y = startingPoint.Y - 0.2188;
              conduitLine1.EndPoint.X = startingPoint.X - 1.25;
              conduitLine1.EndPoint.Y = startingPoint.Y - 1;
              CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
              LineData conduitLine2 = new LineData();
              conduitLine2.Layer = "E-CND1";
              conduitLine2.StartPoint = new SimpleVector3d();
              conduitLine2.EndPoint = new SimpleVector3d();
              conduitLine2.StartPoint.X = startingPoint.X - 1.25;
              conduitLine2.StartPoint.Y = startingPoint.Y - 1 - (5.0 / 16.0);
              conduitLine2.EndPoint.X = startingPoint.X - 1.25;
              conduitLine2.EndPoint.Y = startingPoint.Y - 1.5;
              CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);
              ObjectId meterSymbol = bt["METER CTS (AUTO SINGLE LINE)"];
              using (
                BlockReference acBlkRef = new BlockReference(
                  new Point3d(startingPoint.X - 1.25, startingPoint.Y - 0.5, 0),
                  meterSymbol
                )
              )
              {
                BlockTableRecord acCurSpaceBlkTblRec;
                acCurSpaceBlkTblRec =
                  tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                tr.AddNewlyCreatedDBObject(acBlkRef, true);
              }
              ObjectId breakerSymbol = bt["DS BREAKER (AUTO SINGLE LINE)"];
              using (
                BlockReference acBlkRef = new BlockReference(
                  new Point3d(startingPoint.X - 1.25, startingPoint.Y - 1, 0),
                  breakerSymbol
                )
              )
              {
                BlockTableRecord acCurSpaceBlkTblRec;
                acCurSpaceBlkTblRec =
                  tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                tr.AddNewlyCreatedDBObject(acBlkRef, true);
              }
            }
            else
            {
              LineData conduitLine1 = new LineData();
              conduitLine1.Layer = "E-CND1";
              conduitLine1.StartPoint = new SimpleVector3d();
              conduitLine1.EndPoint = new SimpleVector3d();
              conduitLine1.StartPoint.X = startingPoint.X - 1.25;
              conduitLine1.StartPoint.Y = startingPoint.Y - 0.2188;
              conduitLine1.EndPoint.X = startingPoint.X - 1.25;
              conduitLine1.EndPoint.Y = startingPoint.Y - 0.5;
              CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
              LineData conduitLine2 = new LineData();
              conduitLine2.Layer = "E-CND1";
              conduitLine2.StartPoint = new SimpleVector3d();
              conduitLine2.EndPoint = new SimpleVector3d();
              conduitLine2.StartPoint.X = startingPoint.X - 1.25;
              conduitLine2.StartPoint.Y = startingPoint.Y - 0.75;
              conduitLine2.EndPoint.X = startingPoint.X - 1.25;
              conduitLine2.EndPoint.Y = startingPoint.Y - 1;
              CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);
              LineData conduitLine3 = new LineData();
              conduitLine3.Layer = "E-CND1";
              conduitLine3.StartPoint = new SimpleVector3d();
              conduitLine3.EndPoint = new SimpleVector3d();
              conduitLine3.StartPoint.X = startingPoint.X - 1.25;
              conduitLine3.StartPoint.Y = startingPoint.Y - 1 - (5.0 / 16.0);
              conduitLine3.EndPoint.X = startingPoint.X - 1.25;
              conduitLine3.EndPoint.Y = startingPoint.Y - 1.5;
              CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine3, 1);
              ObjectId meterSymbol = bt["METER (AUTO SINGLE LINE)"];
              using (
                BlockReference acBlkRef = new BlockReference(
                  new Point3d(startingPoint.X - 1.25, startingPoint.Y - 0.625, 0),
                  meterSymbol
                )
              )
              {
                BlockTableRecord acCurSpaceBlkTblRec;
                acCurSpaceBlkTblRec =
                  tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                tr.AddNewlyCreatedDBObject(acBlkRef, true);
              }
              ObjectId breakerSymbol = bt["DS BREAKER (AUTO SINGLE LINE)"];
              using (
                BlockReference acBlkRef = new BlockReference(
                  new Point3d(startingPoint.X - 1.25, startingPoint.Y - 1, 0),
                  breakerSymbol
                )
              )
              {
                BlockTableRecord acCurSpaceBlkTblRec;
                acCurSpaceBlkTblRec =
                  tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                tr.AddNewlyCreatedDBObject(acBlkRef, true);
              }
            }
            if (hasGfp)
            {
              ObjectId gfpSymbol = bt["GFP (AUTO SINGLE LINE)"];
              using (
                BlockReference acBlkRef = new BlockReference(
                  new Point3d(startingPoint.X - 1.25, startingPoint.Y - 1.4375, 0),
                  gfpSymbol
                )
              )
              {
                BlockTableRecord acCurSpaceBlkTblRec;
                acCurSpaceBlkTblRec =
                  tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                tr.AddNewlyCreatedDBObject(acBlkRef, true);
              }
            }
          }
          else
          {
            LineData conduitLine1 = new LineData();
            conduitLine1.Layer = "E-CND1";
            conduitLine1.StartPoint = new SimpleVector3d();
            conduitLine1.EndPoint = new SimpleVector3d();
            conduitLine1.StartPoint.X = startingPoint.X - 1.25;
            conduitLine1.StartPoint.Y = startingPoint.Y - 0.2188;
            conduitLine1.EndPoint.X = startingPoint.X - 1.25;
            conduitLine1.EndPoint.Y = startingPoint.Y - 0.75;
            CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
            LineData conduitLine2 = new LineData();
            conduitLine2.Layer = "E-CND1";
            conduitLine2.StartPoint = new SimpleVector3d();
            conduitLine2.EndPoint = new SimpleVector3d();
            conduitLine2.StartPoint.X = startingPoint.X - 1.25;
            conduitLine2.StartPoint.Y = startingPoint.Y - 0.75 - (5.0 / 16.0);
            conduitLine2.EndPoint.X = startingPoint.X - 1.25;
            conduitLine2.EndPoint.Y = startingPoint.Y - 1.5;
            CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);
            ObjectId breakerSymbol = bt["DS BREAKER (AUTO SINGLE LINE)"];
            using (
              BlockReference acBlkRef = new BlockReference(
                new Point3d(startingPoint.X - 1.25, startingPoint.Y - 0.75, 0),
                breakerSymbol
              )
            )
            {
              BlockTableRecord acCurSpaceBlkTblRec;
              acCurSpaceBlkTblRec =
                tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
              acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
              tr.AddNewlyCreatedDBObject(acBlkRef, true);
            }
            if (hasGfp)
            {
              ObjectId gfpSymbol = bt["GFP (AUTO SINGLE LINE)"];
              using (
                BlockReference acBlkRef = new BlockReference(
                  new Point3d(startingPoint.X - 1.25, startingPoint.Y - 1.1875, 0),
                  gfpSymbol
                )
              )
              {
                BlockTableRecord acCurSpaceBlkTblRec;
                acCurSpaceBlkTblRec =
                  tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                tr.AddNewlyCreatedDBObject(acBlkRef, true);
              }
            }
          }
        }
        else if (!isDistribution && distributionBreakerSize > 0)
        {
          // panel is coming from a distribution section
          if (hasMeter)
          {
            if (hasCts)
            {
              // line to breaker
              LineData lineData1 = new LineData();
              lineData1.Layer = "E-CND1";
              lineData1.StartPoint = new SimpleVector3d();
              lineData1.EndPoint = new SimpleVector3d();
              lineData1.StartPoint.X = startingPoint.X;
              lineData1.StartPoint.Y = startingPoint.Y;
              lineData1.EndPoint.X = startingPoint.X;
              lineData1.EndPoint.Y = startingPoint.Y - (9.0 / 8.0);
              CADObjectCommands.CreateLine(new Point3d(), tr, btr, lineData1, 1);

              ObjectId meterSymbol = bt["METER CTS (AUTO SINGLE LINE)"];
              using (
                BlockReference acBlkRef = new BlockReference(
                  new Point3d(startingPoint.X, startingPoint.Y - 0.5, 0),
                  meterSymbol
                )
              )
              {
                BlockTableRecord acCurSpaceBlkTblRec;
                acCurSpaceBlkTblRec =
                  tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                tr.AddNewlyCreatedDBObject(acBlkRef, true);
              }
            }
            else
            {
              LineData lineData1 = new LineData();
              lineData1.Layer = "E-CND1";
              lineData1.StartPoint = new SimpleVector3d();
              lineData1.EndPoint = new SimpleVector3d();
              lineData1.StartPoint.X = startingPoint.X;
              lineData1.StartPoint.Y = startingPoint.Y;
              lineData1.EndPoint.X = startingPoint.X;
              lineData1.EndPoint.Y = startingPoint.Y - 0.3980;
              CADObjectCommands.CreateLine(new Point3d(), tr, btr, lineData1, 1);
              ObjectId meterSymbol = bt["METER (AUTO SINGLE LINE)"];
              using (
                BlockReference acBlkRef = new BlockReference(
                  new Point3d(startingPoint.X, startingPoint.Y - 0.5230, 0),
                  meterSymbol
                )
              )
              {
                BlockTableRecord acCurSpaceBlkTblRec;
                acCurSpaceBlkTblRec =
                  tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                tr.AddNewlyCreatedDBObject(acBlkRef, true);
              }
              LineData lineData2 = new LineData();
              lineData2.Layer = "E-CND1";
              lineData2.StartPoint = new SimpleVector3d();
              lineData2.EndPoint = new SimpleVector3d();
              lineData2.StartPoint.X = startingPoint.X;
              lineData2.StartPoint.Y = startingPoint.Y - 0.6480;
              lineData2.EndPoint.X = endingPoint.X;
              lineData2.EndPoint.Y = startingPoint.Y - (9.0 / 8.0);
              CADObjectCommands.CreateLine(new Point3d(), tr, btr, lineData2, 1);
            }
          }
          else
          {
            LineData lineData1 = new LineData();
            lineData1.Layer = "E-CND1";
            lineData1.StartPoint = new SimpleVector3d();
            lineData1.EndPoint = new SimpleVector3d();
            lineData1.StartPoint.X = startingPoint.X;
            lineData1.StartPoint.Y = startingPoint.Y;
            lineData1.EndPoint.X = startingPoint.X;
            lineData1.EndPoint.Y = startingPoint.Y - (9.0 / 8.0);
            CADObjectCommands.CreateLine(new Point3d(), tr, btr, lineData1, 1);
          }

          ObjectId breakerSymbol = bt["DS BREAKER (AUTO SINGLE LINE)"];
          using (
            BlockReference acBlkRef = new BlockReference(
              new Point3d(startingPoint.X, startingPoint.Y - 1.125, 0),
              breakerSymbol
            )
          )
          {
            BlockTableRecord acCurSpaceBlkTblRec;
            acCurSpaceBlkTblRec =
              tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
            acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
            tr.AddNewlyCreatedDBObject(acBlkRef, true);
          }

          // line from breaker
          LineData lineData3 = new LineData();
          lineData3.Layer = "E-CND1";
          lineData3.StartPoint = new SimpleVector3d();
          lineData3.EndPoint = new SimpleVector3d();
          lineData3.StartPoint.X = startingPoint.X;
          lineData3.StartPoint.Y = startingPoint.Y - (9.0 / 8.0) - (5.0 / 16.0);
          lineData3.EndPoint.X = endingPoint.X;
          lineData3.EndPoint.Y = endingPoint.Y;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, lineData3, 1);

          // panel rectangle
          Polyline2dData polyData = new Polyline2dData();
          polyData.Layer = "E-SYMBOL";
          polyData.Vertices.Add(new SimpleVector3d(endingPoint.X - (5.0 / 16.0), endingPoint.Y, 0));
          polyData.Vertices.Add(
            new SimpleVector3d(endingPoint.X - (5.0 / 16.0), endingPoint.Y - (17.0 / 16.0), 0)
          );
          polyData.Vertices.Add(
            new SimpleVector3d(endingPoint.X + (5.0 / 16.0), endingPoint.Y - (17.0 / 16.0), 0)
          );
          polyData.Vertices.Add(new SimpleVector3d(endingPoint.X + (5.0 / 16.0), endingPoint.Y, 0));
          polyData.Vertices.Add(new SimpleVector3d(endingPoint.X - (5.0 / 16.0), endingPoint.Y, 0));
          polyData.Closed = true;
          CADObjectCommands.CreatePolyline2d(new Point3d(), tr, btr, polyData, 1);

          LineData lineData4 = new LineData();
          lineData4.Layer = "E-SYM1";
          lineData4.StartPoint = new SimpleVector3d();
          lineData4.EndPoint = new SimpleVector3d();
          lineData4.StartPoint.X = startingPoint.X + (width / 2.0);
          lineData4.StartPoint.Y = startingPoint.Y + 0.25;
          lineData4.EndPoint.X = startingPoint.X + (width / 2.0);
          lineData4.EndPoint.Y = startingPoint.Y + 0.25 - 2;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, lineData4, 1, "HIDDEN");

          if (parentDistance >= 25)
          {
            // main breaker arc
            ArcData arcData2 = new ArcData();
            arcData2.Layer = "E-CND1";
            arcData2.Center = new SimpleVector3d();
            arcData2.Radius = 1.0 / 8.0;
            arcData2.Center.X = endingPoint.X - 0.0302;
            arcData2.Center.Y = endingPoint.Y - (1.0 / 8.0) + 0.0037;
            arcData2.StartAngle = 4.92183;
            arcData2.EndAngle = 1.32645;
            CADObjectCommands.CreateArc(new Point3d(), tr, btr, arcData2, 1);
          }
        }
        else { }
        tr.Commit();
      }
      foreach (var child in children)
      {
        child.Make();
      }
    }
  }

  public class SLTransformer : SingleLine
  {
    public SLTransformer(string id, string name)
    {
      type = "transformer";
      width = 0;
      this.name = name;
      this.id = id;
    }
  }

  public class SLDisconnect : SingleLine
  {
    public SLDisconnect(string name)
    {
      type = "disconnect";
      width = 0;
      this.name = name;
      this.id = id;
    }
  }

  public class SLConduit : SingleLine
  {
    bool isVertical = true;
    string conduitSize,
      wireSize,
      distance;

    public SLConduit(
      bool isVertical,
      string conduitSize,
      string wireSize,
      string distance,
      string name = ""
    )
    {
      type = "conduit";
      width = 1;
      this.name = name;
      this.isVertical = isVertical;
      this.conduitSize = conduitSize;
      this.wireSize = wireSize;
      this.distance = distance;
    }
  }

  public class SLPanelBreaker : SingleLine
  {
    bool isVertical = true;
    string breakerSize;
    string numPoles;

    public SLPanelBreaker(bool isVertical, string breakerSize, string numPoles)
    {
      this.isVertical = isVertical;
      this.breakerSize = breakerSize;
      this.numPoles = numPoles;
      width = 0;
    }
  }
}
