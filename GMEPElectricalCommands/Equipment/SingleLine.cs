using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using DocumentFormat.OpenXml.Presentation;
using GMEPElectricalCommands.GmepDatabase;

namespace ElectricalCommands.Equipment
{
  public class SingleLine
  {
    public double width;
    public string type;
    public string parentType;
    public string name;
    public string id;
    public List<SingleLine> children;
    public Point3d startingPoint;
    public Point3d endingPoint;
    public bool startChildRight;
    public int parentDistance;
    public double aicRating;
    public double parentAicRating;
    public double voltage;
    public string feederWireSize;
    public int feederWireCount;
    public double transformerKva;

    public SingleLine()
    {
      children = new List<SingleLine>();
      parentDistance = 0;
    }

    public virtual double AggregateWidths(bool fromDistribution = false)
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
        child.parentType = type;
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

    public void MakeAicRating(
      Transaction tr,
      BlockTableRecord btr,
      BlockTable bt,
      Database db,
      Point3d endingPoint
    )
    {
      ObjectId aicMarker = bt["AIC MARKER (AUTO SINGLE LINE)"];
      using (
        BlockReference acBlkRef = new BlockReference(
          new Point3d(endingPoint.X, endingPoint.Y + 0.125, 0),
          aicMarker
        )
      )
      {
        BlockTableRecord acCurSpaceBlkTblRec;
        acCurSpaceBlkTblRec =
          tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
        acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
        tr.AddNewlyCreatedDBObject(acBlkRef, true);
      }
      GeneralCommands.CreateAndPositionText(
        tr,
        "~" + Math.Round(aicRating, 0).ToString() + " AIC",
        "gmep",
        0.0938,
        0.85,
        2,
        "E-TXT1",
        new Point3d(endingPoint.X + 0.0678, endingPoint.Y + 0.08, 0)
      );
    }

    public void MakeDistributionCtsMeter(
      Transaction tr,
      BlockTableRecord btr,
      BlockTable bt,
      Database db,
      Point3d startingPoint
    )
    {
      GeneralCommands.CreateAndPositionText(
        tr,
        "(N)",
        "gmep",
        0.0938,
        0.85,
        2,
        "E-TXT1",
        new Point3d(startingPoint.X + 0.36, startingPoint.Y - 0.4, 0)
      );
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

    public void SetFeederWireSizeAndCount(string feederSpec)
    {
      int wireCount = 1;
      if (feederSpec.StartsWith("["))
      {
        wireCount = Int32.Parse(feederSpec[1].ToString());
      }
      feederWireCount = wireCount;
      feederWireSize = Regex.Match(feederSpec, @"(?<=#)([0-9]+(\/0)?( KCMIL)?)").Groups[0].Value;
    }

    public void MakeDistributionMeter(
      Transaction tr,
      BlockTableRecord btr,
      BlockTable bt,
      Database db,
      Point3d startingPoint
    )
    {
      GeneralCommands.CreateAndPositionText(
        tr,
        "(N)",
        "gmep",
        0.0938,
        0.85,
        2,
        "E-TXT1",
        new Point3d(startingPoint.X + 0.13, startingPoint.Y - 0.4, 0)
      );
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

    public void MakeDistributionBreaker(
      Transaction tr,
      BlockTableRecord btr,
      BlockTable bt,
      Database db,
      Point3d startingPoint,
      int mainBreakerSize,
      bool is3Phase
    )
    {
      GeneralCommands.CreateAndPositionText(
        tr,
        "(N)",
        "gmep",
        0.0938,
        0.85,
        2,
        "E-TXT1",
        new Point3d(startingPoint.X + 0.15, startingPoint.Y - 1.19, 0)
      );
      GeneralCommands.CreateAndPositionText(
        tr,
        mainBreakerSize.ToString() + "A",
        "gmep",
        0.0938,
        0.85,
        2,
        "E-TXT1",
        new Point3d(startingPoint.X + 0.15, startingPoint.Y - 1.32, 0)
      );
      GeneralCommands.CreateAndPositionText(
        tr,
        is3Phase ? "3P" : "2P",
        "gmep",
        0.0938,
        0.85,
        2,
        "E-TXT1",
        new Point3d(startingPoint.X + 0.15, startingPoint.Y - 1.45, 0)
      );
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
    }
  }

  public class SLServiceFeeder : SingleLine
  {
    public bool isMultiMeter;
    public string voltageSpec;
    public int amp;

    public SLServiceFeeder(string id, string name, bool isMultiMeter, int amp, string voltageSpec)
    {
      type = "service feeder";
      width = 2.5;
      this.name = name;
      this.id = id;
      this.isMultiMeter = isMultiMeter;
      this.amp = amp;
      this.voltageSpec = voltageSpec;
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
        child.parentType = type;
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
            new Point3d(startingPoint.X + 1.125, startingPoint.Y - 1.85, 0),
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
        GeneralCommands.CreateAndPositionText(
          tr,
          "(N)GND BUS",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(startingPoint.X + 1, startingPoint.Y - 1.75, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "(N)1#3/0 CU.",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(startingPoint.X + 1.6, startingPoint.Y - 2.2, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "TO COLD WATER PIPE",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(startingPoint.X + 1.6, startingPoint.Y - 2.33, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "(N)1#3/0 CU.",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(startingPoint.X + 1.37, startingPoint.Y - 2.6, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "TO COLD BUILDING STEEL",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(startingPoint.X + 1.37, startingPoint.Y - 2.73, 0)
        );
        ObjectId spoon1 = bt["SPOON SMALL LEFT (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(startingPoint.X + 1.35, startingPoint.Y - 2.15, 0),
            spoon1
          )
        )
        {
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        ObjectId spoon2 = bt["SPOON SMALL LEFT (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(startingPoint.X + 1.125, startingPoint.Y - 2.55, 0),
            spoon2
          )
        )
        {
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        ObjectId labelLeader1 = bt["SECTION LABEL LEADER (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(startingPoint.X + 0.25, startingPoint.Y, 0),
            labelLeader1
          )
        )
        {
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        GeneralCommands.CreateAndPositionText(
          tr,
          "(N)" + amp + "A",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(startingPoint.X + 0.35, startingPoint.Y + 0.53, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "UNDERGROUND",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(startingPoint.X + 0.35, startingPoint.Y + 0.38, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "PULL SECTION",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(startingPoint.X + 0.35, startingPoint.Y + 0.25, 0)
        );
        if (isMultiMeter)
        {
          ObjectId labelLeader2 = bt["SECTION LABEL LEADER SMALL (AUTO SINGLE LINE)"];
          using (
            BlockReference acBlkRef = new BlockReference(
              new Point3d(startingPoint.X + 1.625, startingPoint.Y, 0),
              labelLeader2
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
          ObjectId labelLeader2 = bt["SECTION LABEL LEADER (AUTO SINGLE LINE)"];
          using (
            BlockReference acBlkRef = new BlockReference(
              new Point3d(startingPoint.X + 1.625, startingPoint.Y, 0),
              labelLeader2
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
        ObjectId spoon3 = bt["SPOON SMALL RIGHT (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(startingPoint.X + 0.375, startingPoint.Y - 2.25, 0),
            spoon3
          )
        )
        {
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        GeneralCommands.CreateAndPositionText(
          tr,
          "FROM",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(startingPoint.X - 0.25, startingPoint.Y - 2.3, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "SERVICE",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(startingPoint.X - 0.25, startingPoint.Y - 2.43, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "FEEDER",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(startingPoint.X - 0.25, startingPoint.Y - 2.56, 0)
        );
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
    public string conduitSize;
    public string wireSize;
    public string voltageDrop;
    public string voltageSpec;
    public bool is3Phase;

    public SLPanel(string id, string name, bool isDistribution, bool hasMeter, int parentDistance)
    {
      type = "panel";
      width = 2;
      this.name = name;
      this.id = id;
      this.isDistribution = isDistribution;
      this.hasMeter = hasMeter;
      this.parentDistance = parentDistance;
      startChildRight = true;
    }

    public void MakePanel(
      Transaction tr,
      BlockTableRecord btr,
      BlockTable bt,
      Database db,
      Point3d endingPoint,
      string name
    )
    {
      GeneralCommands.CreateAndPositionText(
        tr,
        "(N)",
        "gmep",
        0.0938,
        0.85,
        2,
        "E-TXT1",
        new Point3d(endingPoint.X, endingPoint.Y - 0.44, 0),
        TextHorizontalMode.TextCenter,
        TextVerticalMode.TextBase,
        AttachmentPoint.BaseCenter
      );
      GeneralCommands.CreateAndPositionText(
        tr,
        "PANEL",
        "gmep",
        0.0938,
        0.85,
        2,
        "E-TXT1",
        new Point3d(endingPoint.X, endingPoint.Y - 0.57, 0),
        TextHorizontalMode.TextCenter,
        TextVerticalMode.TextBase,
        AttachmentPoint.BaseCenter
      );
      GeneralCommands.CreateAndPositionText(
        tr,
        "'" + name + "'",
        "gmep",
        0.0938,
        0.85,
        2,
        "E-TXT1",
        new Point3d(endingPoint.X, endingPoint.Y - 0.70, 0),
        TextHorizontalMode.TextCenter,
        TextVerticalMode.TextBase,
        AttachmentPoint.BaseCenter
      ); // panel rectangle
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
    }

    public void MakeMainBreakerArc(
      Transaction tr,
      BlockTableRecord btr,
      BlockTable bt,
      Database db,
      Point3d endingPoint,
      int mainBreakerSize,
      bool is3Phase
    )
    { // main breaker arc
      ArcData arcData2 = new ArcData();
      arcData2.Layer = "E-CND1";
      arcData2.Center = new SimpleVector3d();
      arcData2.Radius = 1.0 / 8.0;
      arcData2.Center.X = endingPoint.X - 0.0302;
      arcData2.Center.Y = endingPoint.Y - (1.0 / 8.0) + 0.0037;
      arcData2.StartAngle = 4.92183;
      arcData2.EndAngle = 1.32645;
      CADObjectCommands.CreateArc(new Point3d(), tr, btr, arcData2, 1);
      ObjectId breakerLeader = bt["BREAKER LEADER RIGHT (AUTO SINGLE LINE)"];
      using (
        BlockReference acBlkRef = new BlockReference(
          new Point3d(endingPoint.X, endingPoint.Y, 0),
          breakerLeader
        )
      )
      {
        BlockTableRecord acCurSpaceBlkTblRec;
        acCurSpaceBlkTblRec =
          tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
        acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
        tr.AddNewlyCreatedDBObject(acBlkRef, true);
      }
      GeneralCommands.CreateAndPositionText(
        tr,
        "(N)" + mainBreakerSize + "A/" + (is3Phase ? "3P" : "2P"),
        "gmep",
        0.0938,
        0.85,
        2,
        "E-TXT1",
        new Point3d(endingPoint.X - 0.42, endingPoint.Y + 0.165, 0),
        TextHorizontalMode.TextCenter,
        TextVerticalMode.TextBase,
        AttachmentPoint.BaseRight
      );
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
              startingPoint.Y - 4.5,
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
          child.parentType = type;
        }
      }
      else
      {
        int index = 0;
        for (int i = 0; i < children.Count; i++)
        {
          SingleLine child = children[i];
          child.parentType = type;
          if (startChildRight)
          {
            if (index == 0)
            {
              child.startChildRight = true;
              child.SetChildEndingPoint(
                new Point3d(endingPoint.X + 2 + (child.children.Count / 2), endingPoint.Y - 3.25, 0)
              );
              child.SetChildStartingPoints(
                new Point3d(endingPoint.X + (5.0 / 16.0), endingPoint.Y - (7.0 / 8.0), 0)
              );
            }
            if (index == 1)
            {
              child.startChildRight = false;
              child.SetChildEndingPoint(
                new Point3d(endingPoint.X - 2 - (child.children.Count / 2), endingPoint.Y - 3.25, 0)
              );
              child.SetChildStartingPoints(
                new Point3d(endingPoint.X - (5.0 / 16.0), endingPoint.Y - (7.0 / 8.0), 0)
              );
            }
            if (index == 2)
            {
              child.startChildRight = true;
              child.SetChildEndingPoint(
                new Point3d(
                  endingPoint.X
                    + (child.children.Count / 2)
                    + (2 * children[i - 2].children.Count)
                    + 4,
                  endingPoint.Y - 3.25,
                  0
                )
              );
              child.SetChildStartingPoints(
                new Point3d(endingPoint.X + (5.0 / 16.0), endingPoint.Y - (2.0 / 8.0), 0)
              );
            }
            if (index == 3)
            {
              child.startChildRight = false;
              child.SetChildEndingPoint(
                new Point3d(
                  endingPoint.X
                    - (child.children.Count / 2)
                    - (2 * children[i - 2].children.Count)
                    - 4,
                  endingPoint.Y - 3.25,
                  0
                )
              );
              child.SetChildStartingPoints(
                new Point3d(endingPoint.X - (5.0 / 16.0), endingPoint.Y - (2.0 / 8.0), 0)
              );
            }
          }
          else
          {
            if (index == 1)
            {
              child.startChildRight = false;
              child.SetChildEndingPoint(new Point3d(endingPoint.X + 2, endingPoint.Y - 3.25, 0));
              child.SetChildStartingPoints(
                new Point3d(endingPoint.X + (5.0 / 16.0), endingPoint.Y - (7.0 / 8.0), 0)
              );
            }
            if (index == 0)
            {
              child.startChildRight = true;
              child.SetChildEndingPoint(new Point3d(endingPoint.X - 2, endingPoint.Y - 3.25, 0));
              child.SetChildStartingPoints(
                new Point3d(endingPoint.X - (5.0 / 16.0), endingPoint.Y - (7.0 / 8.0), 0)
              );
            }
            if (index == 3)
            {
              child.startChildRight = false;
              child.SetChildEndingPoint(
                new Point3d(
                  endingPoint.X
                    + (child.children.Count / 2)
                    + (2 * children[i - 2].children.Count)
                    + 4,
                  endingPoint.Y - 3.25,
                  0
                )
              );
              child.SetChildStartingPoints(
                new Point3d(endingPoint.X + (5.0 / 16.0), endingPoint.Y - (2.0 / 8.0), 0)
              );
            }
            if (index == 2)
            {
              child.startChildRight = true;
              child.SetChildEndingPoint(
                new Point3d(
                  endingPoint.X
                    - (child.children.Count / 2)
                    - (2 * children[i - 2].children.Count)
                    - 4,
                  endingPoint.Y - 3.25,
                  0
                )
              );
              child.SetChildStartingPoints(
                new Point3d(endingPoint.X - (5.0 / 16.0), endingPoint.Y - (2.0 / 8.0), 0)
              );
            }
          }
          index++;
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
          GeneralCommands.CreateAndPositionText(
            tr,
            "(N)" + distributionBreakerSize.ToString() + "A BUS",
            "gmep",
            0.0938,
            0.85,
            2,
            "E-TXT1",
            new Point3d(startingPoint.X + 0.1, startingPoint.Y - 0.145, 0)
          );
          GeneralCommands.CreateAndPositionText(
            tr,
            "(N)" + (hasMeter ? "METER&" : "") + "MAIN",
            "gmep",
            0.0938,
            0.85,
            2,
            "E-TXT1",
            new Point3d(startingPoint.X - 0.75, startingPoint.Y + 0.53, 0)
          );
          GeneralCommands.CreateAndPositionText(
            tr,
            "BREAKER",
            "gmep",
            0.0938,
            0.85,
            2,
            "E-TXT1",
            new Point3d(startingPoint.X - 0.75, startingPoint.Y + 0.38, 0)
          );
          GeneralCommands.CreateAndPositionText(
            tr,
            "SECTION",
            "gmep",
            0.0938,
            0.85,
            2,
            "E-TXT1",
            new Point3d(startingPoint.X - 0.75, startingPoint.Y + 0.25, 0)
          );
          GeneralCommands.CreateAndPositionText(
            tr,
            "(N)" + (hasMeter ? "DISTRIBUTION" : "MULTI-METER") + " SECTION",
            "gmep",
            0.0938,
            0.85,
            2,
            "E-TXT1",
            new Point3d(startingPoint.X + 0.6, startingPoint.Y + 0.53, 0)
          );
          string voltageText = "";
          if (voltageSpec.StartsWith("120/208 3"))
          {
            voltageText = "120/208V-3\u0081-4W";
          }
          if (voltageSpec.StartsWith("120/240 1"))
          {
            voltageText = "120/240V-1\u0081-3W";
          }
          if (voltageSpec.StartsWith("277/480 3"))
          {
            voltageText = "277/480V-3\u0081-4W";
          }
          if (voltageSpec.StartsWith("120/240 3"))
          {
            voltageText = "120/240V-3\u0081-4W";
          }
          GeneralCommands.CreateAndPositionText(
            tr,
            distributionBreakerSize.ToString() + "A " + voltageText,
            "gmep",
            0.0938,
            0.85,
            2,
            "E-TXT1",
            new Point3d(startingPoint.X + 0.6, startingPoint.Y + 0.38, 0)
          );
          GeneralCommands.CreateAndPositionText(
            tr,
            "65 KAIC OR MATCH FAULT CURRENT ON SITE",
            "gmep",
            0.0938,
            0.85,
            2,
            "E-TXT1",
            new Point3d(startingPoint.X + 0.6, startingPoint.Y + 0.25, 0)
          );
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
          ObjectId labelLeader = bt["SECTION LABEL LEADER LONG (AUTO SINGLE LINE)"];
          using (
            BlockReference acBlkRef = new BlockReference(
              new Point3d(startingPoint.X + 0.5, startingPoint.Y, 0),
              labelLeader
            )
          )
          {
            BlockTableRecord acCurSpaceBlkTblRec;
            acCurSpaceBlkTblRec =
              tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
            acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
            tr.AddNewlyCreatedDBObject(acBlkRef, true);
          }
          if (hasMeter)
          {
            GeneralCommands.CreateAndPositionText(
              tr,
              "(N)",
              "gmep",
              0.0938,
              0.85,
              2,
              "E-TXT1",
              new Point3d(startingPoint.X - 1.0725, startingPoint.Y - 1.07, 0)
            );
            GeneralCommands.CreateAndPositionText(
              tr,
              distributionBreakerSize.ToString() + "A",
              "gmep",
              0.0938,
              0.85,
              2,
              "E-TXT1",
              new Point3d(startingPoint.X - 1.0725, startingPoint.Y - 1.20, 0)
            );
            GeneralCommands.CreateAndPositionText(
              tr,
              is3Phase ? "3P" : "2P",
              "gmep",
              0.0938,
              0.85,
              2,
              "E-TXT1",
              new Point3d(startingPoint.X - 1.0725, startingPoint.Y - 1.33, 0)
            );
            if (hasCts)
            {
              GeneralCommands.CreateAndPositionText(
                tr,
                "(N)",
                "gmep",
                0.0938,
                0.85,
                2,
                "E-TXT1",
                new Point3d(startingPoint.X - 0.92, startingPoint.Y - 0.369, 0)
              );
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
              GeneralCommands.CreateAndPositionText(
                tr,
                "(N)",
                "gmep",
                0.0938,
                0.85,
                2,
                "E-TXT1",
                new Point3d(startingPoint.X - 1.14, startingPoint.Y - 0.51, 0)
              );
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
            GeneralCommands.CreateAndPositionText(
              tr,
              "(N)",
              "gmep",
              0.0938,
              0.85,
              2,
              "E-TXT1",
              new Point3d(startingPoint.X - 1.0725, startingPoint.Y - 0.82, 0)
            );
            GeneralCommands.CreateAndPositionText(
              tr,
              distributionBreakerSize.ToString() + "A",
              "gmep",
              0.0938,
              0.85,
              2,
              "E-TXT1",
              new Point3d(startingPoint.X - 1.0725, startingPoint.Y - 0.95, 0)
            );
            GeneralCommands.CreateAndPositionText(
              tr,
              is3Phase ? "3P" : "2P",
              "gmep",
              0.0938,
              0.85,
              2,
              "E-TXT1",
              new Point3d(startingPoint.X - 1.0725, startingPoint.Y - 1.08, 0)
            );
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
              MakeDistributionCtsMeter(tr, btr, bt, db, startingPoint);
            }
            else
            {
              MakeDistributionMeter(tr, btr, bt, db, startingPoint);
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
          if (voltageSpec.Contains("3"))
          {
            is3Phase = true;
          }
          MakeDistributionBreaker(tr, btr, bt, db, startingPoint, mainBreakerSize, is3Phase);
          MakePanel(tr, btr, bt, db, endingPoint, name);
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
            MakeMainBreakerArc(tr, btr, bt, db, endingPoint, mainBreakerSize, is3Phase);
          }
          double voltage = 208;
          if (voltageSpec.Contains("480"))
          {
            voltage = 480;
          }
          if (voltageSpec.Contains("240"))
          {
            voltage = 240;
          }
          (
            string firstLine,
            string secondLine,
            string thirdLine,
            string supplemental1,
            string supplemental2,
            string supplemental3
          ) = CADObjectCommands.GetWireAndConduitSizeText(
            mainBreakerSize,
            mainBreakerSize,
            parentDistance + 10,
            voltage,
            1,
            is3Phase ? 3 : 1
          );
          CADObjectCommands.AddWireAndConduitTextToPlan(
            db,
            new Point3d(endingPoint.X, endingPoint.Y + 0.5, 0),
            firstLine,
            secondLine,
            thirdLine,
            supplemental1,
            supplemental2,
            supplemental3,
            false
          );
          SetFeederWireSizeAndCount(firstLine);
          if (parentType == "transformer")
          {
            aicRating = CADObjectCommands.GetAicRatingFromTransformer(
              transformerKva,
              1,
              0.03,
              parentDistance + 10,
              feederWireCount,
              voltage,
              feederWireSize,
              is3Phase
            );
          }
          else
          {
            aicRating = CADObjectCommands.GetAicRating(
              parentAicRating,
              parentDistance + 10,
              feederWireCount,
              voltage,
              feederWireSize,
              is3Phase
            );
          }
          MakeAicRating(tr, btr, bt, db, endingPoint);
        }
        else
        {
          // panel is subpanel
          LineData lineData1 = new LineData();
          lineData1.Layer = "E-CND1";
          lineData1.StartPoint = new SimpleVector3d();
          lineData1.EndPoint = new SimpleVector3d();
          lineData1.StartPoint.X = startingPoint.X;
          lineData1.StartPoint.Y = startingPoint.Y;
          lineData1.EndPoint.X = startingPoint.X + (endingPoint.X - startingPoint.X);
          lineData1.EndPoint.Y = startingPoint.Y;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, lineData1, 1);
          LineData lineData2 = new LineData();
          lineData2.Layer = "E-CND1";
          lineData2.StartPoint = new SimpleVector3d();
          lineData2.EndPoint = new SimpleVector3d();
          lineData2.StartPoint.X = startingPoint.X + (endingPoint.X - startingPoint.X);
          lineData2.StartPoint.Y = startingPoint.Y;
          lineData2.EndPoint.X = endingPoint.X;
          lineData2.EndPoint.Y = endingPoint.Y;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, lineData2, 1);

          if (endingPoint.X > startingPoint.X)
          {
            // right panel breaker
            ArcData arcData1 = new ArcData();
            arcData1.Layer = "E-CND1";
            arcData1.Center = new SimpleVector3d();
            arcData1.Radius = 0.1038;
            arcData1.Center.X = startingPoint.X - 0.1015;
            arcData1.Center.Y = startingPoint.Y - 0.0216;
            arcData1.StartAngle = 0.20944;
            arcData1.EndAngle = 2.89725;
            CADObjectCommands.CreateArc(new Point3d(), tr, btr, arcData1, 1);
            ObjectId breakerLeader = bt["BREAKER LEADER LEFT (AUTO SINGLE LINE)"];
            using (
              BlockReference acBlkRef = new BlockReference(
                new Point3d(startingPoint.X, startingPoint.Y, 0),
                breakerLeader
              )
            )
            {
              BlockTableRecord acCurSpaceBlkTblRec;
              acCurSpaceBlkTblRec =
                tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
              acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
              tr.AddNewlyCreatedDBObject(acBlkRef, true);
            }
            GeneralCommands.CreateAndPositionText(
              tr,
              "(N)" + mainBreakerSize + "A/" + (is3Phase ? "3P" : "2P"),
              "gmep",
              0.0938,
              0.85,
              2,
              "E-TXT1",
              new Point3d(startingPoint.X + 0.42, startingPoint.Y + 0.165, 0),
              TextHorizontalMode.TextCenter,
              TextVerticalMode.TextBase,
              AttachmentPoint.BaseLeft
            );
          }
          else
          {
            // left panel breaker
            ArcData arcData1 = new ArcData();
            arcData1.Layer = "E-CND1";
            arcData1.Center = new SimpleVector3d();
            arcData1.Radius = 0.1038;
            arcData1.Center.X = startingPoint.X + 0.1015;
            arcData1.Center.Y = startingPoint.Y - 0.0216;
            arcData1.StartAngle = 0.20944;
            arcData1.EndAngle = 2.89725;
            CADObjectCommands.CreateArc(new Point3d(), tr, btr, arcData1, 1);
            ObjectId breakerLeader = bt["BREAKER LEADER RIGHT (AUTO SINGLE LINE)"];
            using (
              BlockReference acBlkRef = new BlockReference(
                new Point3d(startingPoint.X, startingPoint.Y, 0),
                breakerLeader
              )
            )
            {
              BlockTableRecord acCurSpaceBlkTblRec;
              acCurSpaceBlkTblRec =
                tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
              acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
              tr.AddNewlyCreatedDBObject(acBlkRef, true);
            }
            GeneralCommands.CreateAndPositionText(
              tr,
              "(N)" + mainBreakerSize + "A/" + (is3Phase ? "3P" : "2P"),
              "gmep",
              0.0938,
              0.85,
              2,
              "E-TXT1",
              new Point3d(startingPoint.X - 0.42, startingPoint.Y + 0.165, 0),
              TextHorizontalMode.TextCenter,
              TextVerticalMode.TextBase,
              AttachmentPoint.BaseRight
            );
          }
          if (parentDistance >= 25)
          {
            // main breaker arc
            MakeMainBreakerArc(tr, btr, bt, db, endingPoint, mainBreakerSize, is3Phase);
          }
          double voltage = 208;
          if (voltageSpec.Contains("480"))
          {
            voltage = 480;
          }
          if (voltageSpec.Contains("240"))
          {
            voltage = 240;
          }
          (
            string firstLine,
            string secondLine,
            string thirdLine,
            string supplemental1,
            string supplemental2,
            string supplemental3
          ) = CADObjectCommands.GetWireAndConduitSizeText(
            mainBreakerSize,
            mainBreakerSize,
            parentDistance + 10,
            voltage,
            1,
            is3Phase ? 3 : 1
          );
          CADObjectCommands.AddWireAndConduitTextToPlan(
            db,
            new Point3d(endingPoint.X, endingPoint.Y + 0.5, 0),
            firstLine,
            secondLine,
            thirdLine,
            supplemental1,
            supplemental2,
            supplemental3,
            false
          );
          SetFeederWireSizeAndCount(firstLine);
          if (parentType == "transformer")
          {
            aicRating = CADObjectCommands.GetAicRatingFromTransformer(
              transformerKva,
              1,
              0.03,
              parentDistance + 10,
              feederWireCount,
              voltage,
              feederWireSize,
              is3Phase
            );
          }
          else
          {
            aicRating = CADObjectCommands.GetAicRating(
              parentAicRating,
              parentDistance + 10,
              feederWireCount,
              voltage,
              feederWireSize,
              is3Phase
            );
          }
          MakePanel(tr, btr, bt, db, endingPoint, name);
          MakeAicRating(tr, btr, bt, db, endingPoint);
        }
        tr.Commit();
      }
      foreach (var child in children)
      {
        child.parentAicRating = aicRating;
        child.Make();
      }
    }
  }

  public class SLTransformer : SingleLine
  {
    public int distributionBreakerSize;
    public bool hasMeter;
    public bool hasCts;
    public bool is3Phase;
    public string voltageSpec;
    public int mainBreakerSize;
    public string grounding;
    public double kva;

    public SLTransformer(string id, string name)
    {
      type = "transformer";
      width = 2;
      this.name = name;
      this.id = id;
    }

    public override void SetChildStartingPoints(Point3d startingPoint)
    {
      this.startingPoint = startingPoint;
      if (children.Count > 0)
      {
        children[0]
          .SetChildEndingPoint(new Point3d(endingPoint.X, endingPoint.Y - 0.3739 - 2.5, 0));
        children[0].SetChildStartingPoints(new Point3d(endingPoint.X, endingPoint.Y - 0.3739, 0));
        children[0].parentType = type;
        children[0].transformerKva = kva;
      }
    }

    public void MakeTransformer(
      Transaction tr,
      BlockTableRecord btr,
      BlockTable bt,
      Database db,
      Point3d endingPoint
    )
    {
      ObjectId discSymbol = bt["TRANSFORMER (AUTO SINGLE LINE)"];
      using (
        BlockReference acBlkRef = new BlockReference(
          new Point3d(endingPoint.X, endingPoint.Y, 0),
          discSymbol
        )
      )
      {
        BlockTableRecord acCurSpaceBlkTblRec;
        acCurSpaceBlkTblRec =
          tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
        acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
        tr.AddNewlyCreatedDBObject(acBlkRef, true);
      }
      GeneralCommands.CreateAndPositionText(
        tr,
        grounding,
        "gmep",
        0.0938,
        0.85,
        2,
        "E-TXT1",
        new Point3d(endingPoint.X - 0.25, endingPoint.Y - 0.86, 0),
        TextHorizontalMode.TextCenter,
        TextVerticalMode.TextBase,
        AttachmentPoint.BaseRight
      );
      GeneralCommands.CreateAndPositionText(
        tr,
        "TO BUILDING STEEL",
        "gmep",
        0.0938,
        0.85,
        2,
        "E-TXT1",
        new Point3d(endingPoint.X - 0.25, endingPoint.Y - 1.0, 0),
        TextHorizontalMode.TextCenter,
        TextVerticalMode.TextBase,
        AttachmentPoint.BaseRight
      );
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
        if (distributionBreakerSize > 0)
        {
          if (hasMeter)
          {
            if (hasCts)
            {
              MakeDistributionCtsMeter(tr, btr, bt, db, startingPoint);
            }
            else
            {
              MakeDistributionMeter(tr, btr, bt, db, startingPoint);
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
            lineData1.EndPoint.X = endingPoint.X;
            lineData1.EndPoint.Y = endingPoint.Y;
            CADObjectCommands.CreateLine(new Point3d(), tr, btr, lineData1, 1);
          }
          if (voltageSpec.Contains("3"))
          {
            is3Phase = true;
          }
          LineData lineData4 = new LineData();
          lineData4.Layer = "E-SYM1";
          lineData4.StartPoint = new SimpleVector3d();
          lineData4.EndPoint = new SimpleVector3d();
          lineData4.StartPoint.X = startingPoint.X + (width / 2.0);
          lineData4.StartPoint.Y = startingPoint.Y + 0.25;
          lineData4.EndPoint.X = startingPoint.X + (width / 2.0);
          lineData4.EndPoint.Y = startingPoint.Y + 0.25 - 2;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, lineData4, 1, "HIDDEN");
        }
        else
        {
          LineData lineData1 = new LineData();
          lineData1.Layer = "E-CND1";
          lineData1.StartPoint = new SimpleVector3d();
          lineData1.EndPoint = new SimpleVector3d();
          lineData1.StartPoint.X = startingPoint.X;
          lineData1.StartPoint.Y = startingPoint.Y;
          lineData1.EndPoint.X = endingPoint.X;
          lineData1.EndPoint.Y = endingPoint.Y;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, lineData1, 1);
        }
        (
          string firstLine,
          string secondLine,
          string thirdLine,
          string supplemental1,
          string supplemental2,
          string supplemental3
        ) = CADObjectCommands.GetWireAndConduitSizeText(
          mainBreakerSize,
          mainBreakerSize,
          parentDistance + 10,
          voltage,
          1,
          is3Phase ? 3 : 1
        );
        CADObjectCommands.AddWireAndConduitTextToPlan(
          db,
          new Point3d(endingPoint.X, endingPoint.Y + 0.5, 0),
          firstLine,
          secondLine,
          thirdLine,
          supplemental1,
          supplemental2,
          supplemental3,
          false
        );
        SetFeederWireSizeAndCount(firstLine);
        MakeTransformer(tr, btr, bt, db, endingPoint);
        aicRating = CADObjectCommands.GetAicRating(
          parentAicRating,
          parentDistance + 10,
          feederWireCount,
          voltage,
          feederWireSize,
          is3Phase
        );
        MakeAicRating(tr, btr, bt, db, endingPoint);
        tr.Commit();
      }
      foreach (var child in children)
      {
        child.parentAicRating = aicRating;
        child.Make();
      }
    }
  }

  public class SLDisconnect : SingleLine
  {
    public bool fromDistribution;
    public bool hasMeter;
    public bool hasCts;
    public int mainBreakerSize;
    public bool is3Phase;
    public double voltage;

    public SLDisconnect(string name)
    {
      type = "disconnect";
      width = 0;
      this.name = name;
    }

    public override void SetChildStartingPoints(Point3d startingPoint)
    {
      this.startingPoint = startingPoint;
      children[0].SetChildEndingPoint(new Point3d(endingPoint.X, endingPoint.Y - 0.1201 - 2.5, 0));
      children[0].SetChildStartingPoints(new Point3d(endingPoint.X, endingPoint.Y - 0.1201, 0));
      children[0].parentType = type;
    }

    public void MakeDisconnect(
      Transaction tr,
      BlockTableRecord btr,
      BlockTable bt,
      Database db,
      Point3d endingPoint
    )
    {
      ObjectId discSymbol = bt["DISCONNECT (AUTO SINGLE LINE)"];
      using (
        BlockReference acBlkRef = new BlockReference(
          new Point3d(endingPoint.X, endingPoint.Y, 0),
          discSymbol
        )
      )
      {
        BlockTableRecord acCurSpaceBlkTblRec;
        acCurSpaceBlkTblRec =
          tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
        acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
        tr.AddNewlyCreatedDBObject(acBlkRef, true);
      }

      ObjectId arrowSymbol = bt["RIGHT ARROW (AUTO SINGLE LINE)"];
      using (
        BlockReference acBlkRef = new BlockReference(
          new Point3d(endingPoint.X - 0.0601, endingPoint.Y - 0.0601, 0),
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
      string text = "(N)";
      switch (mainBreakerSize)
      {
        case var _ when mainBreakerSize <= 30:
          text += "30AS/";
          break;
        case var _ when mainBreakerSize <= 60:
          text += "60AS/";
          break;
        case var _ when mainBreakerSize <= 100:
          text += "100AS/";
          break;
        case var _ when mainBreakerSize <= 200:
          text += "200AS/";
          break;
        case var _ when mainBreakerSize <= 400:
          text += "400AS/";
          break;
        case var _ when mainBreakerSize <= 600:
          text += "600AS/";
          break;
      }
      text += is3Phase ? "3P" : "2P";
      GeneralCommands.CreateAndPositionText(
        tr,
        text,
        "gmep",
        0.0938,
        0.85,
        2,
        "E-TXT1",
        new Point3d(endingPoint.X - 0.25, endingPoint.Y - 0.037, 0),
        TextHorizontalMode.TextCenter,
        TextVerticalMode.TextBase,
        AttachmentPoint.BaseRight
      );
      GeneralCommands.CreateAndPositionText(
        tr,
        "DISCONNECT",
        "gmep",
        0.0938,
        0.85,
        2,
        "E-TXT1",
        new Point3d(endingPoint.X - 0.25, endingPoint.Y - 0.18, 0),
        TextHorizontalMode.TextCenter,
        TextVerticalMode.TextBase,
        AttachmentPoint.BaseRight
      );
      GeneralCommands.CreateAndPositionText(
        tr,
        "FOR XFMR '" + name + "'",
        "gmep",
        0.0938,
        0.85,
        2,
        "E-TXT1",
        new Point3d(endingPoint.X - 0.25, endingPoint.Y - 0.31, 0),
        TextHorizontalMode.TextCenter,
        TextVerticalMode.TextBase,
        AttachmentPoint.BaseRight
      );
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
        if (fromDistribution)
        {
          if (hasMeter)
          {
            if (hasCts)
            {
              MakeDistributionCtsMeter(tr, btr, bt, db, startingPoint);
            }
            else
            {
              MakeDistributionMeter(tr, btr, bt, db, startingPoint);
            }
          }
          else
          {
            LineData lineData2 = new LineData();
            lineData2.Layer = "E-CND1";
            lineData2.StartPoint = new SimpleVector3d();
            lineData2.EndPoint = new SimpleVector3d();
            lineData2.StartPoint.X = startingPoint.X;
            lineData2.StartPoint.Y = startingPoint.Y;
            lineData2.EndPoint.X = endingPoint.X;
            lineData2.EndPoint.Y = startingPoint.Y - (9.0 / 8.0);
            CADObjectCommands.CreateLine(new Point3d(), tr, btr, lineData2, 1);
          }
          MakeDistributionBreaker(tr, btr, bt, db, startingPoint, mainBreakerSize, is3Phase);
          LineData lineData3 = new LineData();
          lineData3.Layer = "E-CND1";
          lineData3.StartPoint = new SimpleVector3d();
          lineData3.EndPoint = new SimpleVector3d();
          lineData3.StartPoint.X = startingPoint.X;
          lineData3.StartPoint.Y = startingPoint.Y - (9.0 / 8.0) - (5.0 / 16.0);
          lineData3.EndPoint.X = endingPoint.X;
          lineData3.EndPoint.Y = endingPoint.Y;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, lineData3, 1);
          LineData lineData4 = new LineData();
          lineData4.Layer = "E-SYM1";
          lineData4.StartPoint = new SimpleVector3d();
          lineData4.EndPoint = new SimpleVector3d();
          lineData4.StartPoint.X = startingPoint.X + (width / 2.0);
          lineData4.StartPoint.Y = startingPoint.Y + 0.25;
          lineData4.EndPoint.X = startingPoint.X + (width / 2.0);
          lineData4.EndPoint.Y = startingPoint.Y + 0.25 - 2;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, lineData4, 1, "HIDDEN");
        }
        else { }
        (
          string firstLine,
          string secondLine,
          string thirdLine,
          string supplemental1,
          string supplemental2,
          string supplemental3
        ) = CADObjectCommands.GetWireAndConduitSizeText(
          mainBreakerSize,
          mainBreakerSize,
          parentDistance + 10,
          voltage,
          1,
          is3Phase ? 3 : 1
        );
        CADObjectCommands.AddWireAndConduitTextToPlan(
          db,
          new Point3d(endingPoint.X, endingPoint.Y + 0.5, 0),
          firstLine,
          secondLine,
          thirdLine,
          supplemental1,
          supplemental2,
          supplemental3,
          false
        );
        SetFeederWireSizeAndCount(firstLine);
        aicRating = CADObjectCommands.GetAicRating(
          parentAicRating,
          parentDistance + 10,
          feederWireCount,
          voltage,
          feederWireSize,
          is3Phase
        );
        MakeAicRating(tr, btr, bt, db, endingPoint);
        MakeDisconnect(tr, btr, bt, db, endingPoint);
        tr.Commit();
      }
      foreach (var child in children)
      {
        child.parentAicRating = aicRating;
        child.Make();
      }
    }
  }
}
