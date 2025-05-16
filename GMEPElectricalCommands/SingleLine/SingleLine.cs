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

namespace ElectricalCommands.SingleLine
{
  public enum Status
  {
    New,
    Existing,
    Relocated,
  }

  public enum NodeType
  {
    Undefined,
    Service,
    Meter,
    MainBreaker,
    DistributionBus,
    DistributionBreaker,
    Panel,
    PanelBreaker,
    Transformer,
    Disconnect,
    Equipment,
  }

  public class SingleLine
  {
    public double AicRating;

    public void MakeAicRating(
      Transaction tr,
      BlockTableRecord btr,
      BlockTable bt,
      Database db,
      Point3d EndingPoint
    )
    {
      ObjectId aicMarker = bt["AIC MARKER (AUTO SINGLE LINE)"];
      using (
        BlockReference acBlkRef = new BlockReference(
          new Point3d(EndingPoint.X, EndingPoint.Y + 0.125, 0),
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
        "~" + Math.Round(AicRating, 0).ToString() + " AIC",
        "gmep",
        0.0938,
        0.85,
        2,
        "E-TXT1",
        new Point3d(EndingPoint.X + 0.0678, EndingPoint.Y + 0.08, 0)
      );
    }

    public static string GetFeederWireSizeAndCount(string feederSpec)
    {
      int wireCount = 1;
      if (feederSpec.StartsWith("["))
      {
        wireCount = Int32.Parse(feederSpec[1].ToString());
      }
      int feederWireCount = wireCount;
      return Regex.Match(feederSpec, @"(?<=#)([0-9]+(\/0)?( KCMIL)?)").Groups[0].Value;
    }

    public static void MakeDistributionBreakerCombo(
      ElectricalEntity.DistributionBreaker distributionBreaker,
      Point3d currentPoint,
      bool childIsExisting
    )
    {
      MakeDistributionBreaker(
        distributionBreaker,
        new Point3d(currentPoint.X, currentPoint.Y - 0.75, 0)
      );
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);

        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.ColorIndex = distributionBreaker.IsExisting() ? 8 : 256;
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X;
        conduitLine1.StartPoint.Y = currentPoint.Y;
        conduitLine1.EndPoint.X = currentPoint.X;
        conduitLine1.EndPoint.Y = currentPoint.Y - 0.75;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
        LineData conduitLine2 = new LineData();
        conduitLine2.Layer = "E-CND1";
        conduitLine2.ColorIndex = childIsExisting ? 8 : 256;
        conduitLine2.StartPoint = new SimpleVector3d();
        conduitLine2.EndPoint = new SimpleVector3d();
        conduitLine2.StartPoint.X = currentPoint.X;
        conduitLine2.StartPoint.Y = currentPoint.Y - 0.75 - (5.0 / 16.0);
        conduitLine2.EndPoint.X = currentPoint.X;
        conduitLine2.EndPoint.Y = currentPoint.Y - 1.6875;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);
        tr.Commit();
      }
    }

    public static void MakeDistributionCtsMeterCombo(
      ElectricalEntity.Meter meter,
      Point3d currentPoint
    )
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.ColorIndex = meter.IsExisting() ? 8 : 256;
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X;
        conduitLine1.StartPoint.Y = currentPoint.Y;
        conduitLine1.EndPoint.X = currentPoint.X;
        conduitLine1.EndPoint.Y = currentPoint.Y - 1.6875;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);

        tr.Commit();
      }
      MakeCtsMeter(meter, currentPoint);
    }

    public static void MakeDistributionMeterCombo(
      ElectricalEntity.Meter meter,
      Point3d currentPoint,
      bool childIsExisting
    )
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.ColorIndex = meter.IsExisting() ? 8 : 256;
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X;
        conduitLine1.StartPoint.Y = currentPoint.Y;
        conduitLine1.EndPoint.X = currentPoint.X;
        conduitLine1.EndPoint.Y = currentPoint.Y - 0.5;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
        LineData conduitLine2 = new LineData();
        conduitLine2.Layer = "E-CND1";
        conduitLine2.ColorIndex = childIsExisting ? 8 : 256;
        conduitLine2.StartPoint = new SimpleVector3d();
        conduitLine2.EndPoint = new SimpleVector3d();
        conduitLine2.StartPoint.X = currentPoint.X;
        conduitLine2.StartPoint.Y = currentPoint.Y - 0.75;
        conduitLine2.EndPoint.X = currentPoint.X;
        conduitLine2.EndPoint.Y = currentPoint.Y - 1.6875;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);
        tr.Commit();
      }
      MakeMeter(meter, new Point3d(currentPoint.X, currentPoint.Y - 0.625, 0));
    }

    public static void MakeDistributionCtsMeterAndBreakerCombo(
      ElectricalEntity.Meter meter,
      ElectricalEntity.DistributionBreaker distributionBreaker,
      Point3d currentPoint,
      bool childIsExisting
    )
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.ColorIndex = meter.IsExisting() ? 8 : 256;
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X;
        conduitLine1.StartPoint.Y = currentPoint.Y;
        conduitLine1.EndPoint.X = currentPoint.X;
        conduitLine1.EndPoint.Y = currentPoint.Y - 1;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
        LineData conduitLine2 = new LineData();
        conduitLine2.Layer = "E-CND1";
        conduitLine2.ColorIndex = childIsExisting ? 8 : 256;
        conduitLine2.StartPoint = new SimpleVector3d();
        conduitLine2.EndPoint = new SimpleVector3d();
        conduitLine2.StartPoint.X = currentPoint.X;
        conduitLine2.StartPoint.Y = currentPoint.Y - 1 - (5.0 / 16.0);
        conduitLine2.EndPoint.X = currentPoint.X;
        conduitLine2.EndPoint.Y = currentPoint.Y - 1.6875;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);
        tr.Commit();
      }
      MakeCtsMeter(meter, new Point3d(currentPoint.X, currentPoint.Y - 0.5, 0));
      MakeDistributionBreaker(
        distributionBreaker,
        new Point3d(currentPoint.X, currentPoint.Y - 1, 0)
      );
    }

    public static void MakeDistributionMeterAndBreakerCombo(
      ElectricalEntity.Meter meter,
      ElectricalEntity.DistributionBreaker distributionBreaker,
      Point3d currentPoint,
      bool childIsExisting
    )
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.ColorIndex = meter.IsExisting() ? 8 : 256;
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X;
        conduitLine1.StartPoint.Y = currentPoint.Y;
        conduitLine1.EndPoint.X = currentPoint.X;
        conduitLine1.EndPoint.Y = currentPoint.Y - 0.5;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
        LineData conduitLine2 = new LineData();
        conduitLine2.Layer = "E-CND1";
        conduitLine2.ColorIndex = distributionBreaker.IsExisting() ? 8 : 256;
        conduitLine2.StartPoint = new SimpleVector3d();
        conduitLine2.EndPoint = new SimpleVector3d();
        conduitLine2.StartPoint.X = currentPoint.X;
        conduitLine2.StartPoint.Y = currentPoint.Y - 0.75;
        conduitLine2.EndPoint.X = currentPoint.X;
        conduitLine2.EndPoint.Y = currentPoint.Y - 1;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);
        LineData conduitLine3 = new LineData();
        conduitLine3.Layer = "E-CND1";
        conduitLine3.ColorIndex = childIsExisting ? 8 : 256;
        conduitLine3.StartPoint = new SimpleVector3d();
        conduitLine3.EndPoint = new SimpleVector3d();
        conduitLine3.StartPoint.X = currentPoint.X;
        conduitLine3.StartPoint.Y = currentPoint.Y - 1 - (5.0 / 16.0);
        conduitLine3.EndPoint.X = currentPoint.X;
        conduitLine3.EndPoint.Y = currentPoint.Y - 1.6875;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine3, 1);
        tr.Commit();
      }
      MakeMeter(meter, new Point3d(currentPoint.X, currentPoint.Y - 0.625, 0));
      MakeDistributionBreaker(
        distributionBreaker,
        new Point3d(currentPoint.X, currentPoint.Y - 1, 0)
      );
    }

    public static void MakeMainBreaker(
      ElectricalEntity.MainBreaker mainBreaker,
      Point3d currentPoint
    )
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        ObjectId breakerSymbol = bt["DS BREAKER (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X, currentPoint.Y, 0),
            breakerSymbol
          )
        )
        {
          acBlkRef.Layer = mainBreaker.IsExisting() ? "E-SYM-EXISTING" : "E-SYM1";
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        if (mainBreaker.AmpRating >= 1000 && mainBreaker.LineVoltage > 277)
        {
          ObjectId gfpSymbol = bt["GFP (AUTO SINGLE LINE)"];
          using (
            BlockReference acBlkRef = new BlockReference(
              new Point3d(currentPoint.X, currentPoint.Y - 0.4346, 0),
              gfpSymbol
            )
          )
          {
            acBlkRef.Layer = mainBreaker.IsExisting() ? "E-SYM-EXISTING" : "E-SYM1";
            BlockTableRecord acCurSpaceBlkTblRec;
            acCurSpaceBlkTblRec =
              tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
            acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
            tr.AddNewlyCreatedDBObject(acBlkRef, true);
          }
        }
        GeneralCommands.CreateAndPositionText(
          tr,
          (mainBreaker.GetStatusAbbr()),
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.1775, currentPoint.Y - 0.0578, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          mainBreaker.AmpRating.ToString() + "A",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.1775, currentPoint.Y - 0.1878, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          mainBreaker.NumPoles.ToString() + "P",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.1775, currentPoint.Y - 0.3178, 0)
        );
        tr.Commit();
      }
      if (mainBreaker.IsExisting())
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X + 0.4675, currentPoint.Y + 0.0316, 0));
      }
    }

    public static void MakeDistributionBreaker(
      ElectricalEntity.DistributionBreaker distributionBreaker,
      Point3d currentPoint
    )
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        ObjectId breakerSymbol = bt["DS BREAKER (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X, currentPoint.Y, 0),
            breakerSymbol
          )
        )
        {
          acBlkRef.Layer = distributionBreaker.IsExisting() ? "E-SYM-EXISTING" : "E-SYM1";
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        GeneralCommands.CreateAndPositionText(
          tr,
          (distributionBreaker.GetStatusAbbr()),
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.1775, currentPoint.Y - 0.0578, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          distributionBreaker.AmpRating.ToString() + "A",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.1775, currentPoint.Y - 0.1878, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          distributionBreaker.NumPoles.ToString() + "P",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.1775, currentPoint.Y - 0.3178, 0)
        );
        tr.Commit();
      }
      if (distributionBreaker.IsExisting())
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X + 0.2634, currentPoint.Y - 0.4385, 0));
      }
    }

    public static void MakeCtsMeter(ElectricalEntity.Meter meter, Point3d currentPoint)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        GeneralCommands.CreateAndPositionText(
          tr,
          (meter.GetStatusAbbr()),
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.33, currentPoint.Y + 0.1611, 0)
        );
        string blockName = "METER CTS (AUTO SINGLE LINE)";
        if (meter.IsSpace)
        {
          blockName = "METER CTS SPACE (AUTO SINGLE LINE)";
        }
        ObjectId meterSymbol = bt[blockName];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X, currentPoint.Y, 0),
            meterSymbol
          )
        )
        {
          acBlkRef.Layer = meter.IsExisting() ? "E-SYM-EXISTING" : "E-SYM1";
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        tr.Commit();
      }
      if (meter.IsExisting())
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X + 0.5, currentPoint.Y, 0));
      }
    }

    public static void MakeMeter(ElectricalEntity.Meter meter, Point3d currentPoint)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        GeneralCommands.CreateAndPositionText(
          tr,
          (meter.GetStatusAbbr()),
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.1441, currentPoint.Y + 0.115, 0)
        );
        string blockName = "METER (AUTO SINGLE LINE)";
        if (meter.IsSpace)
        {
          blockName = "METER SPACE (AUTO SINGLE LINE)";
        }
        ObjectId meterSymbol = bt[blockName];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X, currentPoint.Y, 0),
            meterSymbol
          )
        )
        {
          acBlkRef.Layer = meter.IsExisting() ? "E-SYM-EXISTING" : "E-SYM1";
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        tr.Commit();
      }
      if (meter.IsExisting())
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X + 0.2595, currentPoint.Y - 0.0250, 0));
      }
    }

    public static void MakePanel(ElectricalEntity.Panel panel, Point3d currentPoint)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        GeneralCommands.CreateAndPositionText(
          tr,
          (panel.GetStatusAbbr()),
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X, currentPoint.Y - 0.44, 0),
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
          new Point3d(currentPoint.X, currentPoint.Y - 0.57, 0),
          TextHorizontalMode.TextCenter,
          TextVerticalMode.TextBase,
          AttachmentPoint.BaseCenter
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "'" + panel.Name.ToUpper().Replace("PANEL", "").Trim() + "'",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X, currentPoint.Y - 0.70, 0),
          TextHorizontalMode.TextCenter,
          TextVerticalMode.TextBase,
          AttachmentPoint.BaseCenter
        ); // panel rectangle
        Polyline2dData polyData = new Polyline2dData();
        polyData.Layer = "E-SYMBOL";
        polyData.ColorIndex =
          panel.IsExisting() && panel.Kva == 0 && panel.LoadAmperage == 0 ? 8 : 256;
        polyData.Vertices.Add(new SimpleVector3d(currentPoint.X - (5.0 / 16.0), currentPoint.Y, 0));
        polyData.Vertices.Add(
          new SimpleVector3d(currentPoint.X - (5.0 / 16.0), currentPoint.Y - (17.0 / 16.0), 0)
        );
        polyData.Vertices.Add(
          new SimpleVector3d(currentPoint.X + (5.0 / 16.0), currentPoint.Y - (17.0 / 16.0), 0)
        );
        polyData.Vertices.Add(new SimpleVector3d(currentPoint.X + (5.0 / 16.0), currentPoint.Y, 0));
        polyData.Closed = true;
        CADObjectCommands.CreatePolyline2d(new Point3d(), tr, btr, polyData, 1);

        if (!panel.IsMlo)
        {
          // Make main breaker
          ArcData arcData = new ArcData();
          arcData.Layer = "E-CND1";
          arcData.ColorIndex = panel.IsExisting() ? 8 : 256;
          arcData.Center = new SimpleVector3d();
          arcData.Radius = 1.0 / 8.0;
          arcData.Center.X = currentPoint.X - 0.0302;
          arcData.Center.Y = currentPoint.Y - (1.0 / 8.0) + 0.0037;
          arcData.StartAngle = 4.92183;
          arcData.EndAngle = 1.32645;
          CADObjectCommands.CreateArc(new Point3d(), tr, btr, arcData, 1);
          ObjectId breakerLeader = bt["BREAKER LEADER RIGHT (AUTO SINGLE LINE)"];
          using (
            BlockReference acBlkRef = new BlockReference(
              new Point3d(currentPoint.X, currentPoint.Y, 0),
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
            "(N)" + panel.MainAmpRating + "A/" + (panel.Voltage.Contains("3") ? "3P" : "2P"),
            "gmep",
            0.0938,
            0.85,
            2,
            "E-TXT1",
            new Point3d(currentPoint.X - 0.42, currentPoint.Y + 0.165, 0),
            TextHorizontalMode.TextCenter,
            TextVerticalMode.TextBase,
            AttachmentPoint.BaseRight
          );
        }
        if (panel.Kva == 0 && panel.LoadAmperage == 0 && panel.IsExisting())
        {
          GeneralCommands.CreateAndPositionText(
            tr,
            "NO WORK",
            "gmep",
            0.0938,
            0.85,
            2,
            "E-TXT1",
            new Point3d(currentPoint.X, currentPoint.Y - 1.1826, 0),
            TextHorizontalMode.TextCenter,
            TextVerticalMode.TextBase,
            AttachmentPoint.BaseCenter
          );
          InsertKeyedNoteMarker(1, new Point3d(currentPoint.X, currentPoint.Y - 1.3191, 0));
        }
        else
        {
          GeneralCommands.CreateAndPositionText(
            tr,
            Math.Round(panel.Kva, 0).ToString() + " KVA",
            "gmep",
            0.0938,
            0.85,
            2,
            "E-TXT1",
            new Point3d(currentPoint.X + 0.2874, currentPoint.Y - 1.1826, 0),
            TextHorizontalMode.TextCenter,
            TextVerticalMode.TextBase,
            AttachmentPoint.BaseRight
          );
          GeneralCommands.CreateAndPositionText(
            tr,
            Math.Round(panel.LoadAmperage, 0).ToString() + " A",
            "gmep",
            0.0938,
            0.85,
            2,
            "E-TXT1",
            new Point3d(currentPoint.X + 0.0977, currentPoint.Y - 1.3126, 0),
            TextHorizontalMode.TextCenter,
            TextVerticalMode.TextBase,
            AttachmentPoint.BaseRight
          );
        }

        tr.Commit();
      }
      if (panel.IsExisting())
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X, currentPoint.Y - 0.1541, 0));
      }
    }

    public static void MakeRightPanelBreaker(
      ElectricalEntity.PanelBreaker panelBreaker,
      Point3d currentPoint
    )
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        ArcData arcData1 = new ArcData();
        arcData1.ColorIndex = panelBreaker.IsExisting() ? 8 : 256;
        arcData1.Layer = "E-CND1";
        arcData1.Center = new SimpleVector3d();
        arcData1.Radius = 0.1038;
        arcData1.Center.X = currentPoint.X - 0.1015;
        arcData1.Center.Y = currentPoint.Y - 0.0216;
        arcData1.StartAngle = 0.20944;
        arcData1.EndAngle = 2.89725;
        CADObjectCommands.CreateArc(new Point3d(), tr, btr, arcData1, 1);
        ObjectId breakerLeader = bt["BREAKER LEADER LEFT (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X, currentPoint.Y, 0),
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
          (panelBreaker.GetStatusAbbr())
            + panelBreaker.AmpRating.ToString()
            + "A/"
            + panelBreaker.NumPoles.ToString()
            + "P",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.42, currentPoint.Y + 0.165, 0),
          TextHorizontalMode.TextCenter,
          TextVerticalMode.TextBase,
          AttachmentPoint.BaseLeft
        );
        tr.Commit();
      }
      if (panelBreaker.IsExisting())
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X + 0.7, currentPoint.Y + 0.4050, 0));
      }
    }

    public static void MakeLeftPanelBreaker(
      ElectricalEntity.PanelBreaker panelBreaker,
      Point3d currentPoint
    )
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        // left panel breaker
        ArcData arcData1 = new ArcData();
        arcData1.ColorIndex = panelBreaker.IsExisting() ? 8 : 256;
        arcData1.Layer = "E-CND1";
        arcData1.Center = new SimpleVector3d();
        arcData1.Radius = 0.1038;
        arcData1.Center.X = currentPoint.X + 0.1015;
        arcData1.Center.Y = currentPoint.Y - 0.0216;
        arcData1.StartAngle = 0.20944;
        arcData1.EndAngle = 2.89725;
        CADObjectCommands.CreateArc(new Point3d(), tr, btr, arcData1, 1);
        ObjectId breakerLeader = bt["BREAKER LEADER RIGHT (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X, currentPoint.Y, 0),
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
          (panelBreaker.GetStatusAbbr())
            + panelBreaker.AmpRating.ToString()
            + "A/"
            + panelBreaker.NumPoles.ToString()
            + "P",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X - 0.42, currentPoint.Y + 0.165, 0),
          TextHorizontalMode.TextCenter,
          TextVerticalMode.TextBase,
          AttachmentPoint.BaseRight
        );
        tr.Commit();
      }
      if (panelBreaker.IsExisting())
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X - 0.7, currentPoint.Y + 0.4050, 0));
      }
    }

    public static void MakeDisconnect(ElectricalEntity.Disconnect disconnect, Point3d currentPoint)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        ObjectId discSymbol = bt["DISCONNECT (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X, currentPoint.Y, 0),
            discSymbol
          )
        )
        {
          acBlkRef.Layer = disconnect.IsExisting() ? "E-SYM-EXISTING" : "E-SYM1";
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }

        ObjectId arrowSymbol = bt["RIGHT ARROW (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X - 0.0601, currentPoint.Y - 0.0601, 0),
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
        string text = (disconnect.GetStatusAbbr());
        text += disconnect.AsSize.ToString() + "AS/";
        text += disconnect.AfSize.ToString() + "AF/";
        text += disconnect.NumPoles + "P";
        GeneralCommands.CreateAndPositionText(
          tr,
          text,
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X - 0.25, currentPoint.Y - 0.037, 0),
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
          new Point3d(currentPoint.X - 0.25, currentPoint.Y - 0.18, 0),
          TextHorizontalMode.TextCenter,
          TextVerticalMode.TextBase,
          AttachmentPoint.BaseRight
        );
        tr.Commit();
      }
      if (disconnect.IsExisting())
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X - 0.7210, currentPoint.Y + 0.1891, 0));
      }
    }

    public static void MakeTransformer(
      ElectricalEntity.Transformer transformer,
      Point3d currentPoint
    )
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;

      double voltage = 208;
      if (transformer.Voltage.StartsWith("480"))
      {
        voltage = 480;
      }
      string grounding = String.Empty;
      double amperage = transformer.Kva * 1000 / voltage;
      switch (amperage)
      {
        case var _ when amperage <= 100:
          grounding += "3/4\"C. (1#8CU.)";
          break;
        case var _ when amperage <= 125:
          grounding += "3/4\"C. (1#8CU.)";
          break;
        case var _ when amperage <= 150:
          grounding += "3/4\"C. (1#8CU.)";
          break;
        case var _ when amperage <= 175:
          grounding += "3/4\"C. (1#8CU.)";
          break;
        case var _ when amperage <= 200:
          grounding += "3/4\"C. (1#6CU.)";
          break;
        case var _ when amperage <= 225:
          grounding += "3/4\"C. (1#6CU.)";
          break;
        case var _ when amperage <= 250:
          grounding += "3/4\"C. (1#6CU.)";
          break;
        case var _ when amperage <= 275:
          grounding += "3/4\"C. (1#6CU.)";
          break;
        case var _ when amperage <= 400:
          grounding += "3/4\"C. (1#3CU.)";
          break;
        case var _ when amperage <= 500:
          grounding += "3/4\"C. (1#2CU.)";
          break;
        case var _ when amperage <= 600:
          grounding += "3/4\"C. (1#1CU.)";
          break;
        case var _ when amperage <= 800:
          grounding += "3/4\"C. (1#1/0CU.)";
          break;
      }
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        ObjectId discSymbol = bt["TRANSFORMER (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X, currentPoint.Y, 0),
            discSymbol
          )
        )
        {
          acBlkRef.Layer = transformer.IsExisting() ? "E-SYM-EXISTING" : "E-SYM1";
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        GeneralCommands.CreateAndPositionText(
          tr,
          transformer.GetStatusAbbr() + grounding,
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X - 0.25, currentPoint.Y - 0.86, 0),
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
          new Point3d(currentPoint.X - 0.25, currentPoint.Y - 1.0, 0),
          TextHorizontalMode.TextCenter,
          TextVerticalMode.TextBase,
          AttachmentPoint.BaseRight
        );
        ObjectId arrowSymbol = bt["RIGHT ARROW LONG (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X - 0.0601 - 0.1273, currentPoint.Y - 0.0601, 0),
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
        string line1 =
          transformer.GetStatusAbbr()
          + transformer.Name.Replace(transformer.Voltage, "").Replace(", ", "").ToUpper();
        string line2 = transformer.Voltage;
        string line3 = $"Z=3.5";
        GeneralCommands.CreateAndPositionText(
          tr,
          line1,
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X - 0.3333, currentPoint.Y - 0.05, 0),
          TextHorizontalMode.TextCenter,
          TextVerticalMode.TextBase,
          AttachmentPoint.BaseRight
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          line2,
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X - 0.3333, currentPoint.Y - 0.18, 0),
          TextHorizontalMode.TextCenter,
          TextVerticalMode.TextBase,
          AttachmentPoint.BaseRight
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          line3,
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X - 0.3333, currentPoint.Y - 0.31, 0),
          TextHorizontalMode.TextCenter,
          TextVerticalMode.TextBase,
          AttachmentPoint.BaseRight
        );
        tr.Commit();
      }
      if (transformer.IsExisting())
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X - 1.0573, currentPoint.Y + 0.1649, 0));
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X - 1.0573, currentPoint.Y - 0.6340, 0));
      }
    }

    public static void MakeDistributionChildConduit(Point3d currentPoint, bool existing)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.ColorIndex = existing ? 8 : 256;
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X;
        conduitLine1.StartPoint.Y = currentPoint.Y;
        conduitLine1.EndPoint.X = currentPoint.X;
        conduitLine1.EndPoint.Y = currentPoint.Y - 2.5;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
        tr.Commit();
      }
      if (existing)
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X - 0.4437, currentPoint.Y - 1.6585, 0));
      }
    }

    public static Point3d MakeConduitFromTransformer(Point3d currentPoint, bool existing)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      double xOffset = 0;
      double yOffset = -2.5;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.ColorIndex = existing ? 8 : 256;
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X;
        conduitLine1.StartPoint.Y = currentPoint.Y;
        conduitLine1.EndPoint.X = currentPoint.X + xOffset;
        conduitLine1.EndPoint.Y = currentPoint.Y + yOffset;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
        tr.Commit();
      }
      return new Point3d(currentPoint.X + xOffset, currentPoint.Y + yOffset, 0);
    }

    public static Point3d MakeConduitFromDisconnect(Point3d currentPoint, bool existing)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      double xOffset = 0;
      double yOffset = -2.5;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.ColorIndex = existing ? 8 : 256;
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X;
        conduitLine1.StartPoint.Y = currentPoint.Y;
        conduitLine1.EndPoint.X = currentPoint.X + xOffset;
        conduitLine1.EndPoint.Y = currentPoint.Y + yOffset;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
        tr.Commit();
      }
      if (existing)
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X - 0.4437, currentPoint.Y - 1.6585, 0));
      }
      return new Point3d(currentPoint.X + xOffset, currentPoint.Y + yOffset, 0);
    }

    public static Point3d MakePanelChildConduit(int index, Point3d currentPoint, bool existing)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      double xOffset = 1.5;
      double yOffset = -2.5;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        if (index == 0)
        {
          if (existing)
            InsertKeyedNoteMarker(
              1,
              new Point3d(currentPoint.X + 1.0581, currentPoint.Y - 1.6585, 0)
            );
        }
        if (index == 1)
        {
          xOffset = xOffset * -1;
          if (existing)
            InsertKeyedNoteMarker(
              1,
              new Point3d(currentPoint.X - 1.9562, currentPoint.Y - 1.6585, 0)
            );
        }
        if (index == 2)
        {
          xOffset = xOffset * 2.25;
          yOffset = yOffset - 0.7;
          if (existing)
            InsertKeyedNoteMarker(
              1,
              new Point3d(currentPoint.X + 2.9147, currentPoint.Y - 2.3895, 0)
            );
        }
        if (index == 3)
        {
          xOffset = xOffset * -2.25;
          yOffset = yOffset - 0.7;
          if (existing)
            InsertKeyedNoteMarker(
              1,
              new Point3d(currentPoint.X - 3.8321, currentPoint.Y - 2.3895, 0)
            );
        }

        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.ColorIndex = existing ? 8 : 256;
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X;
        conduitLine1.StartPoint.Y = currentPoint.Y;
        conduitLine1.EndPoint.X = currentPoint.X + xOffset;
        conduitLine1.EndPoint.Y = currentPoint.Y;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
        LineData conduitLine2 = new LineData();
        conduitLine2.Layer = "E-CND1";
        conduitLine2.ColorIndex = existing ? 8 : 256;
        conduitLine2.StartPoint = new SimpleVector3d();
        conduitLine2.EndPoint = new SimpleVector3d();
        conduitLine2.StartPoint.X = currentPoint.X + xOffset;
        conduitLine2.StartPoint.Y = currentPoint.Y;
        conduitLine2.EndPoint.X = currentPoint.X + xOffset;
        conduitLine2.EndPoint.Y = currentPoint.Y + yOffset;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);
        tr.Commit();
      }
      return new Point3d(currentPoint.X + xOffset, currentPoint.Y + yOffset, 0);
    }

    public static void MakePullSection(ElectricalEntity.Service service, Point3d currentPoint)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        LayerTable layerTable = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
        if (layerTable.Has("E-SYM-EXISTING") == false)
        {
          using (LayerTableRecord layerTableRecord = new LayerTableRecord())
          {
            layerTableRecord.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
              Autodesk.AutoCAD.Colors.ColorMethod.ByAci,
              8
            );
            layerTableRecord.Name = "E-SYM-EXISTING";

            tr.GetObject(db.LayerTableId, OpenMode.ForWrite);
            layerTable.Add(layerTableRecord);
            tr.AddNewlyCreatedDBObject(layerTableRecord, true);
          }
        }
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.ColorIndex = service.IsExisting() ? 8 : 256;
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X + 0.375;
        conduitLine1.StartPoint.Y = currentPoint.Y - 0.2813;
        conduitLine1.EndPoint.X = currentPoint.X + 0.375;
        conduitLine1.EndPoint.Y = currentPoint.Y - 2;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);

        LineData conduitLine2 = new LineData();
        conduitLine2.Layer = "E-CND1";
        conduitLine2.ColorIndex = service.IsExisting() ? 8 : 256;
        conduitLine2.StartPoint = new SimpleVector3d();
        conduitLine2.EndPoint = new SimpleVector3d();
        conduitLine2.StartPoint.X = currentPoint.X + 0.375;
        conduitLine2.StartPoint.Y = currentPoint.Y - 0.2813;
        conduitLine2.EndPoint.X = currentPoint.X + 0.75;
        conduitLine2.EndPoint.Y = currentPoint.Y - 0.2813;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);

        LineData feederLine = new LineData();
        feederLine.Layer = "E-CND1";
        feederLine.ColorIndex = service.IsExisting() ? 8 : 256;
        feederLine.StartPoint = new SimpleVector3d();
        feederLine.EndPoint = new SimpleVector3d();
        feederLine.StartPoint.X = currentPoint.X + 0.375;
        feederLine.StartPoint.Y = currentPoint.Y - 2;
        feederLine.EndPoint.X = currentPoint.X + 0.375;
        feederLine.EndPoint.Y = currentPoint.Y - 2.875;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, feederLine, 1, "HIDDEN2");
        ObjectId arrowSymbol = bt["DOWN ARROW (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X + 0.375, currentPoint.Y - 2.875, 0),
            arrowSymbol
          )
        )
        {
          acBlkRef.Layer = service.IsExisting() ? "E-SYM-EXISTING" : "E-SYM1";
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        ObjectId dash1 = bt["SERVICE FEEDER DASH (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X + 0.375, currentPoint.Y - 0.75, 0),
            dash1
          )
        )
        {
          acBlkRef.Layer = service.IsExisting() ? "E-SYM-EXISTING" : "E-SYM1";
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        ObjectId dash2 = bt["SERVICE FEEDER DASH (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X + 0.375, currentPoint.Y - 1.75, 0),
            dash2
          )
        )
        {
          acBlkRef.Layer = service.IsExisting() ? "E-SYM-EXISTING" : "E-SYM1";
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }

        ObjectId spoon = bt["SPOON SMALL RIGHT (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X + 0.375, currentPoint.Y - 2.25, 0),
            spoon
          )
        )
        {
          acBlkRef.Layer = service.IsExisting() ? "E-SYM-EXISTING" : "E-SYM1";
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
          new Point3d(currentPoint.X - 0.25, currentPoint.Y - 2.3, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "SERVICE",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X - 0.25, currentPoint.Y - 2.43, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "FEEDER",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X - 0.25, currentPoint.Y - 2.56, 0)
        );

        GeneralCommands.CreateAndPositionText(
          tr,
          service.GetStatusAbbr() + service.AmpRating.ToString() + "A",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.4, currentPoint.Y + 0.53, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "UNDERGROUND",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.4, currentPoint.Y + 0.38, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "PULL SECTION",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.4, currentPoint.Y + 0.25, 0)
        );

        ObjectId labelLeader1 = bt["SECTION LABEL LEADER (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X + 0.25, currentPoint.Y, 0),
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
        tr.Commit();
      }
      if (service.IsExisting())
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X + 0.4739, currentPoint.Y + 0.76, 0));
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X - 0.1433, currentPoint.Y - 2.6941, 0));
      }
    }

    public static void MakeGroundingBusSection(Point3d currentPoint, bool existing)
    {
      MakeGroundingBus(currentPoint, existing);
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);

        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.ColorIndex = existing ? 8 : 256;
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X;
        conduitLine1.StartPoint.Y = currentPoint.Y - 0.2813;
        conduitLine1.EndPoint.X = currentPoint.X + 1.75;
        conduitLine1.EndPoint.Y = currentPoint.Y - 0.2813;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);

        tr.Commit();
      }
    }

    public static void MakeMainSection(Point3d currentPoint, bool existing)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);

        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.ColorIndex = existing ? 8 : 256;
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X + 1.25;
        conduitLine1.StartPoint.Y = currentPoint.Y - 0.2813;
        conduitLine1.EndPoint.X = currentPoint.X + 1.75;
        conduitLine1.EndPoint.Y = currentPoint.Y - 0.2813;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);

        LineData conduitLine2 = new LineData();
        conduitLine2.Layer = "E-CND1";
        conduitLine2.ColorIndex = existing ? 8 : 256;
        conduitLine2.StartPoint = new SimpleVector3d();
        conduitLine2.EndPoint = new SimpleVector3d();
        conduitLine2.StartPoint.X = currentPoint.X;
        conduitLine2.StartPoint.Y = currentPoint.Y - 0.2813;
        conduitLine2.EndPoint.X = currentPoint.X + 0.5;
        conduitLine2.EndPoint.Y = currentPoint.Y - 0.2813;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);

        LineData conduitLine3 = new LineData();
        conduitLine3.Layer = "E-CND1";
        conduitLine3.ColorIndex = existing ? 8 : 256;
        conduitLine3.StartPoint = new SimpleVector3d();
        conduitLine3.EndPoint = new SimpleVector3d();
        conduitLine3.StartPoint.X = currentPoint.X + 0.5;
        conduitLine3.StartPoint.Y = currentPoint.Y - 1.5;
        conduitLine3.EndPoint.X = currentPoint.X + 1.25;
        conduitLine3.EndPoint.Y = currentPoint.Y - 1.5;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine3, 1);

        LineData conduitLine4 = new LineData();
        conduitLine4.Layer = "E-CND1";
        conduitLine4.ColorIndex = existing ? 8 : 256;
        conduitLine4.StartPoint = new SimpleVector3d();
        conduitLine4.EndPoint = new SimpleVector3d();
        conduitLine4.StartPoint.X = currentPoint.X + 1.25;
        conduitLine4.StartPoint.Y = currentPoint.Y - 1.5;
        conduitLine4.EndPoint.X = currentPoint.X + 1.25;
        conduitLine4.EndPoint.Y = currentPoint.Y - 0.2813;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine4, 1);

        ObjectId labelLeader1 = bt["SECTION LABEL LEADER (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X + 0.875, currentPoint.Y, 0),
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

        tr.Commit();
      }
    }

    public static void MakeMainMeterSection(ElectricalEntity.Meter meter, Point3d currentPoint)
    {
      MakeGroundingBus(currentPoint, meter.IsExisting());
      MakeMainSection(currentPoint, meter.IsExisting());
      MakeMeter(meter, new Point3d(currentPoint.X + 0.5, currentPoint.Y - 0.8907, 0));
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.ColorIndex = meter.IsExisting() ? 8 : 256;
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X + 0.5;
        conduitLine1.StartPoint.Y = currentPoint.Y - 0.7657;
        conduitLine1.EndPoint.X = currentPoint.X + 0.5;
        conduitLine1.EndPoint.Y = currentPoint.Y - 0.2813;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);

        LineData conduitLine2 = new LineData();
        conduitLine2.Layer = "E-CND1";
        conduitLine2.ColorIndex = meter.IsExisting() ? 8 : 256;
        conduitLine2.StartPoint = new SimpleVector3d();
        conduitLine2.EndPoint = new SimpleVector3d();
        conduitLine2.StartPoint.X = currentPoint.X + 0.5;
        conduitLine2.StartPoint.Y = currentPoint.Y - 1.0156;
        conduitLine2.EndPoint.X = currentPoint.X + 0.5;
        conduitLine2.EndPoint.Y = currentPoint.Y - 1.5;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);

        GeneralCommands.CreateAndPositionText(
          tr,
          meter.GetStatusAbbr() + "METER",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 1, currentPoint.Y + 0.53, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "SECTION",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 1, currentPoint.Y + 0.38, 0)
        );

        tr.Commit();
      }
      if (meter.IsExisting())
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X + 1.0646, currentPoint.Y + 0.7600, 0));
      }
    }

    public static void MakeMainBreakerSection(
      ElectricalEntity.MainBreaker mainBreaker,
      Point3d currentPoint
    )
    {
      MakeGroundingBus(currentPoint, mainBreaker.IsExisting());
      MakeMainSection(currentPoint, mainBreaker.IsExisting());
      MakeMainBreaker(mainBreaker, new Point3d(currentPoint.X + 0.5, currentPoint.Y - 0.7636, 0));
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.ColorIndex = mainBreaker.IsExisting() ? 8 : 256;
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X + 0.5;
        conduitLine1.StartPoint.Y = currentPoint.Y - 0.7636;
        conduitLine1.EndPoint.X = currentPoint.X + 0.5;
        conduitLine1.EndPoint.Y = currentPoint.Y - 0.2813;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);

        LineData conduitLine2 = new LineData();
        conduitLine2.Layer = "E-CND1";
        conduitLine2.ColorIndex = mainBreaker.IsExisting() ? 8 : 256;
        conduitLine2.StartPoint = new SimpleVector3d();
        conduitLine2.EndPoint = new SimpleVector3d();
        conduitLine2.StartPoint.X = currentPoint.X + 0.5;
        conduitLine2.StartPoint.Y = currentPoint.Y - 1.0761;
        conduitLine2.EndPoint.X = currentPoint.X + 0.5;
        conduitLine2.EndPoint.Y = currentPoint.Y - 1.5;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);

        GeneralCommands.CreateAndPositionText(
          tr,
          mainBreaker.GetStatusAbbr() + "MAIN",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 1, currentPoint.Y + 0.53, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "BREAKER",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 1, currentPoint.Y + 0.38, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "SECTION",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 1, currentPoint.Y + 0.25, 0)
        );

        tr.Commit();
      }
      if (mainBreaker.IsExisting())
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X + 1.0646, currentPoint.Y + 0.7600, 0));
      }
    }

    public static void MakeMainMeterAndBreakerSection(
      ElectricalEntity.Meter meter,
      ElectricalEntity.MainBreaker mainBreaker,
      Point3d currentPoint
    )
    {
      MakeGroundingBus(currentPoint, meter.IsExisting());
      MakeMainSection(currentPoint, meter.IsExisting());
      if (meter.HasCts)
      {
        MakeCtsMeter(meter, new Point3d(currentPoint.X + 0.5, currentPoint.Y - 0.6238, 0));
      }
      else
      {
        MakeMeter(meter, new Point3d(currentPoint.X + 0.5, currentPoint.Y - 0.6238, 0));
      }
      MakeMainBreaker(mainBreaker, new Point3d(currentPoint.X + 0.5, currentPoint.Y - 1, 0));
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        if (meter.HasCts)
        {
          LineData conduitLine1 = new LineData();
          conduitLine1.Layer = "E-CND1";
          conduitLine1.ColorIndex = meter.IsExisting() ? 8 : 256;
          conduitLine1.StartPoint = new SimpleVector3d();
          conduitLine1.EndPoint = new SimpleVector3d();
          conduitLine1.StartPoint.X = currentPoint.X + 0.5;
          conduitLine1.StartPoint.Y = currentPoint.Y - 1;
          conduitLine1.EndPoint.X = currentPoint.X + 0.5;
          conduitLine1.EndPoint.Y = currentPoint.Y - 0.2813;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
          LineData conduitLine2 = new LineData();
          conduitLine2.Layer = "E-CND1";
          conduitLine2.ColorIndex = mainBreaker.IsExisting() ? 8 : 256;
          conduitLine2.StartPoint = new SimpleVector3d();
          conduitLine2.EndPoint = new SimpleVector3d();
          conduitLine2.StartPoint.X = currentPoint.X + 0.5;
          conduitLine2.StartPoint.Y = currentPoint.Y - 1.3125;
          conduitLine2.EndPoint.X = currentPoint.X + 0.5;
          conduitLine2.EndPoint.Y = currentPoint.Y - 1.5;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);
        }
        else
        {
          LineData conduitLine1 = new LineData();
          conduitLine1.Layer = "E-CND1";
          conduitLine1.ColorIndex = meter.IsExisting() ? 8 : 256;
          conduitLine1.StartPoint = new SimpleVector3d();
          conduitLine1.EndPoint = new SimpleVector3d();
          conduitLine1.StartPoint.X = currentPoint.X + 0.5;
          conduitLine1.StartPoint.Y = currentPoint.Y - 0.4988;
          conduitLine1.EndPoint.X = currentPoint.X + 0.5;
          conduitLine1.EndPoint.Y = currentPoint.Y - 0.2813;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
          LineData conduitLine2 = new LineData();
          conduitLine2.Layer = "E-CND1";
          conduitLine2.ColorIndex = mainBreaker.IsExisting() ? 8 : 256;
          conduitLine2.StartPoint = new SimpleVector3d();
          conduitLine2.EndPoint = new SimpleVector3d();
          conduitLine2.StartPoint.X = currentPoint.X + 0.5;
          conduitLine2.StartPoint.Y = currentPoint.Y - 1.0194;
          conduitLine2.EndPoint.X = currentPoint.X + 0.5;
          conduitLine2.EndPoint.Y = currentPoint.Y - 0.7488;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);
          LineData conduitLine3 = new LineData();
          conduitLine3.Layer = "E-CND1";
          conduitLine3.ColorIndex = mainBreaker.IsExisting() ? 8 : 256;
          conduitLine3.StartPoint = new SimpleVector3d();
          conduitLine3.EndPoint = new SimpleVector3d();
          conduitLine3.StartPoint.X = currentPoint.X + 0.5;
          conduitLine3.StartPoint.Y = currentPoint.Y - 1.3319;
          conduitLine3.EndPoint.X = currentPoint.X + 0.5;
          conduitLine3.EndPoint.Y = currentPoint.Y - 1.5;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine3, 1);
        }

        GeneralCommands.CreateAndPositionText(
          tr,
          mainBreaker.GetStatusAbbr() + "METER&MAIN",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 1, currentPoint.Y + 0.53, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "BREAKER",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 1, currentPoint.Y + 0.38, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          "SECTION",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 1, currentPoint.Y + 0.25, 0)
        );

        tr.Commit();
      }
      if (meter.IsExisting())
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X + 1.0646, currentPoint.Y + 0.7600, 0));
      }
    }

    public static void MakeGroundingBus(Point3d currentPoint, bool existing)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        ObjectId gndBus = bt["GND BUS (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X + 0.3771, currentPoint.Y - 1.85, 0),
            gndBus
          )
        )
        {
          acBlkRef.Layer = existing ? "E-SYM-EXISTING" : "E-SYM1";
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        string statusText = existing ? "(E)" : "(N)";
        GeneralCommands.CreateAndPositionText(
          tr,
          statusText + "GND BUS",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.25, currentPoint.Y - 1.75, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          statusText + (existing ? "" : "1#3/0 CU."),
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.85, currentPoint.Y - 2.2, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          (existing ? "GROUND BUS" : "TO COLD WATER PIPE"),
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.85, currentPoint.Y - 2.33, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          statusText + (existing ? "" : "1#3/0 CU."),
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.62, currentPoint.Y - 2.6, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          (existing ? "GROUND BUS" : "TO COLD WATER PIPE"),
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.62, currentPoint.Y - 2.73, 0)
        );
        ObjectId spoon1 = bt["SPOON SMALL LEFT (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X + 0.6, currentPoint.Y - 2.15, 0),
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
            new Point3d(currentPoint.X + 0.375, currentPoint.Y - 2.55, 0),
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
        tr.Commit();
      }
      if (existing)
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X + 1.0933, currentPoint.Y - 1.7003, 0));
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X + 1.7042, currentPoint.Y + -2.1189, 0));
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X + 1.4716, currentPoint.Y - 2.5207, 0));
      }
      else
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X + 1.7042, currentPoint.Y + -2.1189, 0));
        InsertKeyedNoteMarker(2, new Point3d(currentPoint.X + 1.4716, currentPoint.Y - 2.5207, 0));
      }
    }

    public static void MakeEquipment(ElectricalEntity.Equipment equipment, Point3d currentPoint)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      string topText;
      string bottomText;
      int hyphenIndex = equipment.Name.IndexOf('-');
      if (hyphenIndex == -1)
      {
        topText = equipment.Name;
        bottomText = "-";
      }
      else
      {
        topText = equipment.Name.Substring(0, hyphenIndex);
        bottomText = equipment.Name.Substring(hyphenIndex + 1);
      }
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);

        ObjectId tagBlockId = bt["4CFM"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X - 0.2694, currentPoint.Y - 0.2333, 0),
            tagBlockId
          )
        )
        {
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);

          acBlkRef.Layer = "E-TXT1";
          TextStyleTable textStyleTable = (TextStyleTable)
            tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead);
          ObjectId gmepTextStyleId;
          if (textStyleTable.Has("gmep"))
          {
            gmepTextStyleId = textStyleTable["gmep"];
          }
          else
          {
            ed.WriteMessage("\nText style 'gmep' not found. Using default text style.");
            gmepTextStyleId = doc.Database.Textstyle;
          }
          AttributeDefinition attrDef = new AttributeDefinition();
          attrDef.Position = new Point3d(currentPoint.X, currentPoint.Y - 0.1302, 0);
          attrDef.LockPositionInBlock = false;
          attrDef.Tag = "tag";
          attrDef.IsMTextAttributeDefinition = false;
          attrDef.TextString = topText;
          attrDef.Justify = AttachmentPoint.MiddleCenter;
          attrDef.Visible = true;
          attrDef.Invisible = false;
          attrDef.Constant = false;
          attrDef.Height = 0.0938;
          attrDef.WidthFactor = 0.85;
          attrDef.TextStyleId = gmepTextStyleId;
          attrDef.Layer = "0";

          AttributeReference attrRef = new AttributeReference();
          attrRef.SetAttributeFromBlock(
            attrDef,
            Matrix3d.Displacement(new Vector3d(attrDef.Position.X, attrDef.Position.Y, 0))
          );
          acBlkRef.AttributeCollection.AppendAttribute(attrRef);

          AttributeDefinition attrDef2 = new AttributeDefinition();
          attrDef2.Position = new Point3d(currentPoint.X, currentPoint.Y - 0.3602, 0);
          attrDef2.LockPositionInBlock = false;
          attrDef2.Tag = "number";
          attrDef2.IsMTextAttributeDefinition = false;
          attrDef2.TextString = bottomText;
          attrDef2.Justify = AttachmentPoint.MiddleCenter;
          attrDef2.Visible = true;
          attrDef2.Invisible = false;
          attrDef2.Constant = false;
          attrDef2.Height = 0.0938;
          attrDef2.WidthFactor = 0.85;
          attrDef2.TextStyleId = gmepTextStyleId;
          attrDef2.Layer = "0";

          AttributeReference attrRef2 = new AttributeReference();
          attrRef2.SetAttributeFromBlock(
            attrDef2,
            Matrix3d.Displacement(new Vector3d(attrDef2.Position.X, attrDef2.Position.Y, 0))
          );
          acBlkRef.AttributeCollection.AppendAttribute(attrRef2);
          acBlkRef.Layer = "E-TXT1";
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
          tr.Commit();
        }
      }
    }

    public static void MakeDistributionBus(
      ElectricalEntity.DistributionBus distributionBus,
      bool isMultimeter,
      double width,
      Point3d currentPoint,
      string name
    )
    {
      Point3d busBarPoint = new Point3d(currentPoint.X + 0.25, currentPoint.Y - 0.25, 0);
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

        ObjectId labelLeader = bt["SECTION LABEL LEADER LONG (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X + 0.5, currentPoint.Y, 0),
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
        string voltageText = "120/240-1\u0081-3W";
        if (distributionBus.LineVoltage == 208 && distributionBus.Phase == 3)
        {
          voltageText = "120/208V-3\u0081-4W";
        }
        if (distributionBus.LineVoltage == 480 && distributionBus.Phase == 3)
        {
          voltageText = "277/480V-3\u0081-4W";
        }
        if (distributionBus.LineVoltage == 208 && distributionBus.Phase == 1)
        {
          voltageText = "120/208V-1\u0081-3W";
        }
        if (distributionBus.LineVoltage == 240 && distributionBus.Phase == 3)
        {
          voltageText = "120/240V-3\u0081-4W";
        }

        GeneralCommands.CreateAndPositionText(
          tr,
          distributionBus.GetStatusAbbr()
            + (isMultimeter ? "MULTI-METER" : "DISTRIBUTION")
            + " SECTION"
            + (!String.IsNullOrEmpty(name) ? $" '{name}'" : ""),
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.6, currentPoint.Y + 0.53, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          distributionBus.AmpRating.ToString() + "A " + voltageText,
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.6, currentPoint.Y + 0.38, 0)
        );
        if (!distributionBus.IsExisting())
        {
          GeneralCommands.CreateAndPositionText(
            tr,
            "65 KAIC OR MATCH FAULT CURRENT ON SITE",
            "gmep",
            0.0938,
            0.85,
            2,
            "E-TXT1",
            new Point3d(currentPoint.X + 0.6, currentPoint.Y + 0.25, 0)
          );
        }

        GeneralCommands.CreateAndPositionText(
          tr,
          distributionBus.GetStatusAbbr() + distributionBus.AmpRating.ToString() + "A BUS",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(busBarPoint.X + 0.05, busBarPoint.Y + 0.05, 0)
        );

        Polyline2dData polyData = new Polyline2dData();
        polyData.Layer = "E-CND1";
        polyData.ColorIndex = distributionBus.IsExisting() ? 8 : 256;
        polyData.Vertices.Add(new SimpleVector3d(busBarPoint.X, busBarPoint.Y, 0));
        polyData.Vertices.Add(new SimpleVector3d(busBarPoint.X + width - 0.5, busBarPoint.Y, 0));
        polyData.Vertices.Add(
          new SimpleVector3d(busBarPoint.X + width - 0.5, busBarPoint.Y - 0.0625, 0)
        );
        polyData.Vertices.Add(new SimpleVector3d(busBarPoint.X, busBarPoint.Y - 0.0625, 0));
        polyData.Closed = true;
        CADObjectCommands.CreatePolyline2d(new Point3d(), tr, btr, polyData, 1);
        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.ColorIndex = distributionBus.IsExisting() ? 8 : 256;
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X;
        conduitLine1.StartPoint.Y = currentPoint.Y - 0.2813;
        conduitLine1.EndPoint.X = currentPoint.X + 0.25;
        conduitLine1.EndPoint.Y = currentPoint.Y - 0.2813;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
        tr.Commit();
      }
      if (distributionBus.IsExisting())
      {
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X + 0.6721, currentPoint.Y + 0.7600, 0));
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X + 0.1603, currentPoint.Y - 0.14, 0));
      }
    }

    public static void MakeAicRating(double aicRating, Point3d currentPoint)
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
        ObjectId aicMarker = bt["AIC MARKER (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X, currentPoint.Y + 0.1647, 0),
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
          new Point3d(currentPoint.X + 0.0678, currentPoint.Y + 0.1197, 0)
        );
        tr.Commit();
      }
    }

    public static (string, int) AddConduitSpec(
      ElectricalEntity.PlaceableElectricalEntity entity,
      Point3d currentPoint
    )
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      (
        string firstLine,
        string secondLine,
        string thirdLine,
        string supplemental1,
        string supplemental2,
        string supplemental3
      ) = CADObjectCommands.GetWireAndConduitSizeText(
        entity.LoadAmperage,
        entity.AmpRating,
        entity.ParentDistance,
        entity.LineVoltage,
        3,
        entity.Phase
      );
      if (entity.IsExisting() && entity.LoadAmperage == 0 && entity.Kva == 0)
      {
        int i = thirdLine.IndexOf(";");
        thirdLine = thirdLine.Substring(0, i);
      }
      CADObjectCommands.AddWireAndConduitTextToPlan(
        db,
        new Point3d(currentPoint.X, currentPoint.Y + 0.3333, 0),
        (entity.GetStatusAbbr()) + firstLine,
        secondLine,
        thirdLine,
        supplemental1,
        supplemental2,
        supplemental3,
        false
      );

      int feederWireCount = 1;
      if (firstLine.StartsWith("["))
      {
        feederWireCount = Int32.Parse(firstLine[1].ToString());
      }
      string feederWireSize = Regex
        .Match(firstLine, @"(?<=#)([0-9]+(\/0)?( KCMIL)?)")
        .Groups[0]
        .Value;
      return (feederWireSize, feederWireCount);
    }

    public static void InsertKeyedNoteMarker(int index, Point3d location)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);

        ObjectId keyedNoteBlockId = bt["KEYED NOTE (AUTO SINGLE LINE)"];
        using (BlockReference acBlkRef = new BlockReference(location, keyedNoteBlockId))
        {
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          TextStyleTable textStyleTable = (TextStyleTable)
            tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead);
          ObjectId gmepTextStyleId;
          gmepTextStyleId = textStyleTable["gmep"];

          AttributeDefinition attrDef = new AttributeDefinition();
          attrDef.Position = location;
          attrDef.LockPositionInBlock = true;
          attrDef.Tag = index.ToString();
          attrDef.IsMTextAttributeDefinition = false;
          attrDef.TextString = index.ToString();
          attrDef.Justify = AttachmentPoint.MiddleCenter;
          attrDef.Visible = true;
          attrDef.Invisible = false;
          attrDef.Constant = false;
          attrDef.Height = 0.0938;
          attrDef.WidthFactor = 0.85;
          attrDef.TextStyleId = gmepTextStyleId;
          attrDef.Layer = "0";

          AttributeReference attrRef = new AttributeReference();
          attrRef.SetAttributeFromBlock(attrDef, acBlkRef.BlockTransform);
          acBlkRef.AttributeCollection.AppendAttribute(attrRef);
          acBlkRef.Layer = "E-TEXT";
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        tr.Commit();
      }
    }

    public static void InsertNotes(Point3d currentPoint, bool existing, double height)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        ObjectId generalNotes = bt["GMEP SINGLE-LINE DIAGRAM GENERAL NOTES"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X, currentPoint.Y, 0),
            generalNotes
          )
        )
        {
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        if (existing)
        {
          ObjectId keyedNotes = bt["GMEP SINGLE-LINE DIAGRAM KEYED NOTES - EXISTING"];
          using (
            BlockReference acBlkRef = new BlockReference(
              new Point3d(currentPoint.X, currentPoint.Y - 6.86, 0),
              keyedNotes
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
          ObjectId keyedNotes = bt["GMEP SINGLE-LINE DIAGRAM KEYED NOTES - NEW"];
          using (
            BlockReference acBlkRef = new BlockReference(
              new Point3d(currentPoint.X, currentPoint.Y - 6.86, 0),
              keyedNotes
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
        tr.Commit();
      }
    }

    public static void InsertLabel(Point3d currentPoint, double height)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        ObjectId label = bt["SINGLE LINE LABEL"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X, currentPoint.Y - height - 4, 0),
            label
          )
        )
        {
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
          tr.Commit();
        }
      }
    }

    public static void MakeNoWorkText(Point3d currentPoint)
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
        GeneralCommands.CreateAndPositionText(
          tr,
          "NO WORK",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X, currentPoint.Y, 0),
          TextHorizontalMode.TextCenter,
          TextVerticalMode.TextBase,
          AttachmentPoint.BaseCenter
        );
        InsertKeyedNoteMarker(1, new Point3d(currentPoint.X, currentPoint.Y - 0.13, 0));
        tr.Commit();
      }
    }
  }
}
