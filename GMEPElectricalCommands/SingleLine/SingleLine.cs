using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using DocumentFormat.OpenXml.Drawing.Charts;
using ElectricalCommands.ElectricalEntity;
using GMEPElectricalCommands.GmepDatabase;

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

    private static string GetStatusText(ElectricalEntity.ElectricalEntity electricalEntity)
    {
      return "(" + electricalEntity.Status[0].ToString().ToUpper() + ")";
    }

    public static void MakeDistributionBreakerCombo(
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
      Editor ed = doc.Editor;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        GeneralCommands.CreateAndPositionText(
          tr,
          "(N)",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.1775, currentPoint.Y - 0.82, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          distributionBreaker.AmpRating.ToString() + "A",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.1775, currentPoint.Y - 0.95, 0)
        );
        GeneralCommands.CreateAndPositionText(
          tr,
          distributionBreaker.NumPoles.ToString() + "P",
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.1775, currentPoint.Y - 1.08, 0)
        );
        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X;
        conduitLine1.StartPoint.Y = currentPoint.Y;
        conduitLine1.EndPoint.X = currentPoint.X;
        conduitLine1.EndPoint.Y = currentPoint.Y - 0.75;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
        LineData conduitLine2 = new LineData();
        conduitLine2.Layer = "E-CND1";
        conduitLine2.StartPoint = new SimpleVector3d();
        conduitLine2.EndPoint = new SimpleVector3d();
        conduitLine2.StartPoint.X = currentPoint.X;
        conduitLine2.StartPoint.Y = currentPoint.Y - 0.75 - (5.0 / 16.0);
        conduitLine2.EndPoint.X = currentPoint.X;
        conduitLine2.EndPoint.Y = currentPoint.Y - 1.6875;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);
        ObjectId breakerSymbol = bt["DS BREAKER (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X, currentPoint.Y - 0.75, 0),
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
        tr.Commit();
      }
    }

    public static void MakeDistributionCtsMeterAndBreakerCombo(
      ElectricalEntity.Meter meter,
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
      Editor ed = doc.Editor;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X;
        conduitLine1.StartPoint.Y = currentPoint.Y;
        conduitLine1.EndPoint.X = currentPoint.X;
        conduitLine1.EndPoint.Y = currentPoint.Y - 1;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
        LineData conduitLine2 = new LineData();
        conduitLine2.Layer = "E-CND1";
        conduitLine2.StartPoint = new SimpleVector3d();
        conduitLine2.EndPoint = new SimpleVector3d();
        conduitLine2.StartPoint.X = currentPoint.X;
        conduitLine2.StartPoint.Y = currentPoint.Y - 1 - (5.0 / 16.0);
        conduitLine2.EndPoint.X = currentPoint.X;
        conduitLine2.EndPoint.Y = currentPoint.Y - 1.6875;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);
        tr.Commit();
      }
      MakeCtsMeter(meter, currentPoint);
      MakeDistributionBreaker(
        distributionBreaker,
        new Point3d(currentPoint.X, currentPoint.Y - 1, 0)
      );
    }

    public static void MakeDistributionMeterAndBreakerCombo(
      ElectricalEntity.Meter meter,
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
      Editor ed = doc.Editor;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X;
        conduitLine1.StartPoint.Y = currentPoint.Y;
        conduitLine1.EndPoint.X = currentPoint.X;
        conduitLine1.EndPoint.Y = currentPoint.Y - 0.5;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
        LineData conduitLine2 = new LineData();
        conduitLine2.Layer = "E-CND1";
        conduitLine2.StartPoint = new SimpleVector3d();
        conduitLine2.EndPoint = new SimpleVector3d();
        conduitLine2.StartPoint.X = currentPoint.X;
        conduitLine2.StartPoint.Y = currentPoint.Y - 0.75;
        conduitLine2.EndPoint.X = currentPoint.X;
        conduitLine2.EndPoint.Y = currentPoint.Y - 1;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);
        LineData conduitLine3 = new LineData();
        conduitLine3.Layer = "E-CND1";
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
      Editor ed = doc.Editor;
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
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        GeneralCommands.CreateAndPositionText(
          tr,
          GetStatusText(distributionBreaker),
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
      Editor ed = doc.Editor;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        GeneralCommands.CreateAndPositionText(
          tr,
          GetStatusText(meter),
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.33, currentPoint.Y - 0.369, 0)
        );
        ObjectId meterSymbol = bt["METER CTS (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X, currentPoint.Y - 0.5, 0),
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
        tr.Commit();
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
      Editor ed = doc.Editor;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        GeneralCommands.CreateAndPositionText(
          tr,
          GetStatusText(meter),
          "gmep",
          0.0938,
          0.85,
          2,
          "E-TXT1",
          new Point3d(currentPoint.X + 0.1441, currentPoint.Y + 0.115, 0)
        );

        ObjectId meterSymbol = bt["METER (AUTO SINGLE LINE)"];
        using (
          BlockReference acBlkRef = new BlockReference(
            new Point3d(currentPoint.X, currentPoint.Y, 0),
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
        tr.Commit();
      }
    }
  }
}
