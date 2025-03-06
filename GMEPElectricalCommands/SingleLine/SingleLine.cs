using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Accord.Statistics.Filters;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using DocumentFormat.OpenXml.Drawing;
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
        conduitLine2.EndPoint.Y = currentPoint.Y - 1.6875;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine2, 1);
        tr.Commit();
      }
      MakeMeter(meter, new Point3d(currentPoint.X, currentPoint.Y - 0.625, 0));
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
        string blockName = "METER CTS (AUTO SINGLE LINE)";
        if (meter.IsSpace)
        {
          blockName = "METER CTS SPACE (AUTO SINGLE LINE)";
        }
        ObjectId meterSymbol = bt[blockName];
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
          BlockTableRecord acCurSpaceBlkTblRec;
          acCurSpaceBlkTblRec =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
          tr.AddNewlyCreatedDBObject(acBlkRef, true);
        }
        tr.Commit();
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
          GetStatusText(panel),
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
        polyData.Vertices.Add(new SimpleVector3d(currentPoint.X - (5.0 / 16.0), currentPoint.Y, 0));
        polyData.Vertices.Add(
          new SimpleVector3d(currentPoint.X - (5.0 / 16.0), currentPoint.Y - (17.0 / 16.0), 0)
        );
        polyData.Vertices.Add(
          new SimpleVector3d(currentPoint.X + (5.0 / 16.0), currentPoint.Y - (17.0 / 16.0), 0)
        );
        polyData.Vertices.Add(new SimpleVector3d(currentPoint.X + (5.0 / 16.0), currentPoint.Y, 0));
        polyData.Vertices.Add(new SimpleVector3d(currentPoint.X - (5.0 / 16.0), currentPoint.Y, 0));
        polyData.Closed = true;
        CADObjectCommands.CreatePolyline2d(new Point3d(), tr, btr, polyData, 1);

        if (!panel.IsMlo)
        {
          // Make main breaker
          ArcData arcData = new ArcData();
          arcData.Layer = "E-CND1";
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
        tr.Commit();
      }
    }

    public static void MakeMainBreakerArc(ElectricalEntity.Panel panel, Point3d currentPoint)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      bool is3Phase = panel.Voltage.Contains("3");
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        // main breaker arc
        ArcData arcData2 = new ArcData();
        arcData2.Layer = "E-CND1";
        arcData2.Center = new SimpleVector3d();
        arcData2.Radius = 1.0 / 8.0;
        arcData2.Center.X = currentPoint.X - 0.0302;
        arcData2.Center.Y = currentPoint.Y - (1.0 / 8.0) + 0.0037;
        arcData2.StartAngle = 4.92183;
        arcData2.EndAngle = 1.32645;
        CADObjectCommands.CreateArc(new Point3d(), tr, btr, arcData2, 1);
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
          GetStatusText(panel) + panel.MainAmpRating.ToString() + "A/" + (is3Phase ? "3P" : "2P"),
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
          GetStatusText(panelBreaker)
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
          GetStatusText(panelBreaker)
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
        string text = GetStatusText(disconnect);
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
            new Point3d(currentPoint.X - 0.0601 - 0.1273, currentPoint.Y - 0.0601, 0), // HERE test
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
          $"{GetStatusText(transformer)}{transformer.Kva.ToString()}KVA XFMR '{transformer.Name.ToUpper()}'";
        Console.WriteLine("voltg" + transformer.Voltage);
        string line2 =
          $"{transformer.Voltage}\u0081-{(transformer.Voltage.Contains("3") ? "4W" : "3W")}";
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
    }

    public static void MakeDistributionChildConduit(Point3d currentPoint)
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
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X;
        conduitLine1.StartPoint.Y = currentPoint.Y;
        conduitLine1.EndPoint.X = currentPoint.X;
        conduitLine1.EndPoint.Y = currentPoint.Y - 2.3125;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
        tr.Commit();
      }
    }

    public static Point3d MakeConduitFromTransformer(Point3d currentPoint)
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

    public static Point3d MakeConduitFromDisconnect(Point3d currentPoint)
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

    public static Point3d MakePanelChildConduit(int index, Point3d currentPoint)
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
        if (index == 1)
        {
          xOffset = xOffset * -1;
        }
        if (index == 2)
        {
          xOffset = xOffset * 2;
          yOffset = yOffset + 0.75;
        }
        if (index == 3)
        {
          xOffset = xOffset * -2;
          yOffset = yOffset + 0.75;
        }

        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
        LineData conduitLine1 = new LineData();
        conduitLine1.Layer = "E-CND1";
        conduitLine1.StartPoint = new SimpleVector3d();
        conduitLine1.EndPoint = new SimpleVector3d();
        conduitLine1.StartPoint.X = currentPoint.X;
        conduitLine1.StartPoint.Y = currentPoint.Y;
        conduitLine1.EndPoint.X = currentPoint.X + xOffset;
        conduitLine1.EndPoint.Y = currentPoint.Y;
        CADObjectCommands.CreateLine(new Point3d(), tr, btr, conduitLine1, 1);
        LineData conduitLine2 = new LineData();
        conduitLine2.Layer = "E-CND1";
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
  }
}
