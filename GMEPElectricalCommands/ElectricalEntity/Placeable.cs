﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using ElectricalCommands.SingleLine;

namespace ElectricalCommands.ElectricalEntity
{
  public class PlaceableElectricalEntity : ElectricalEntity
  {
    public string ParentId;
    public string ParentName;
    public int ParentDistance;
    public Point3d Location;
    public bool IsHidden;
    public string BlockName;
    public bool Rotate;

    public bool IsPlaced()
    {
      if (Location.X == 0 && Location.Y == 0)
      {
        return true;
      }
      return false;
    }

    public Point3d? Place()
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
          tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

        //remove any previous marker with the same equipId
        var modelSpace = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
        foreach (ObjectId id in modelSpace)
        {
          try
          {
            BlockReference br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
            if (
              br != null
              && br.IsDynamicBlock
              && br.DynamicBlockReferencePropertyCollection.Count > 0
            )
            {
              DynamicBlockReferencePropertyCollection pc =
                br.DynamicBlockReferencePropertyCollection;
              foreach (DynamicBlockReferenceProperty prop in pc)
              {
                if (prop.PropertyName == "gmep_equip_id" && prop.Value as string == Id)
                {
                  BlockReference eraseBlock = (BlockReference)tr.GetObject(id, OpenMode.ForWrite);
                  eraseBlock.Erase();
                }
              }
            }
          }
          catch { }
        }
        tr.Commit();
      }
      double scale = 12;
      ObjectId blockId;
      if (
        CADObjectCommands.Scale <= 0
        && (CADObjectCommands.IsInModel() || CADObjectCommands.IsInLayoutViewport())
      )
      {
        CADObjectCommands.SetScale();
        if (CADObjectCommands.Scale <= 0)
          return null;
      }
      if (CADObjectCommands.IsInModel() || CADObjectCommands.IsInLayoutViewport())
      {
        scale = CADObjectCommands.Scale;
      }
      Point3d firstClickPoint;
      double rotation = 0;

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

        BlockTableRecord block = (BlockTableRecord)tr.GetObject(bt[BlockName], OpenMode.ForRead);
        BlockJig blockJig = new BlockJig();

        PromptResult blockPromptResult = blockJig.DragMe(block.ObjectId, out firstClickPoint);

        if (blockPromptResult.Status == PromptStatus.OK)
        {
          BlockTableRecord curSpace = (BlockTableRecord)
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

          double scaleFactor = 0.25;
          if (
            NodeType != NodeType.Panel
            && NodeType != NodeType.Transformer
            && NodeType != NodeType.Service
            && NodeType != NodeType.DistributionBus
          )
          {
            scaleFactor = scale;
          }

          BlockReference br = new BlockReference(firstClickPoint, block.ObjectId);

          if (blockPromptResult.Status == PromptStatus.OK)
          {
            BlockTableRecord currentSpace =
              tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
            br.ScaleFactors = new Scale3d(0.25 / scaleFactor);
            br.Layer = "E-SYM1";
          }

          if (Rotate)
          {
            RotateJig rotateJig = new RotateJig(br);
            PromptResult rotatePromptResult = ed.Drag(rotateJig);

            if (rotatePromptResult.Status != PromptStatus.OK)
            {
              return new Point3d();
            }
            rotation = br.Rotation;
          }

          curSpace.AppendEntity(br);

          tr.AddNewlyCreatedDBObject(br, true);
          blockId = br.Id;
        }
        else
        {
          return new Point3d();
        }

        tr.Commit();
      }
      double labelOffsetX = 0;
      double labelOffsetY = 0;

      labelOffsetY = labelOffsetY * 0.25 / scale;
      labelOffsetX = labelOffsetX * 0.25 / scale;
      firstClickPoint = new Point3d(
        firstClickPoint.X + labelOffsetX,
        firstClickPoint.Y + labelOffsetY,
        0
      );

      LabelJig jig = new LabelJig(firstClickPoint);
      PromptResult res = ed.Drag(jig);
      if (res.Status != PromptStatus.OK)
        return null;

      Vector3d direction = jig.endPoint - firstClickPoint;
      double angle = direction.GetAngleTo(Vector3d.XAxis, Vector3d.ZAxis);

      Point3d secondClickPoint = jig.endPoint;
      Point3d thirdClickPoint = Point3d.Origin;
      bool thirdClickOccurred = false;
      using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
      {
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite);
        btr.AppendEntity(jig.line);
        tr.AddNewlyCreatedDBObject(jig.line, true);

        tr.Commit();
      }
      if (angle != 0 && angle != Math.PI)
      {
        DynamicLineJig lineJig = new DynamicLineJig(jig.endPoint, scale);
        res = ed.Drag(lineJig);
        if (res.Status == PromptStatus.OK)
        {
          using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
          {
            BlockTableRecord btr = (BlockTableRecord)
              tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite);
            btr.AppendEntity(lineJig.line);
            tr.AddNewlyCreatedDBObject(lineJig.line, true);

            thirdClickPoint = lineJig.line.EndPoint;
            thirdClickOccurred = true;

            tr.Commit();
          }
        }
      }
      Point3d textAlignmentReferencePoint = thirdClickOccurred ? thirdClickPoint : secondClickPoint;
      Point3d comparisonPoint = thirdClickOccurred ? secondClickPoint : firstClickPoint;
      Point3d labelInsertionPoint;
      if (textAlignmentReferencePoint.X > comparisonPoint.X)
      {
        labelInsertionPoint = new Point3d(
          textAlignmentReferencePoint.X + 14.1197 * 0.25 / scale,
          textAlignmentReferencePoint.Y,
          0
        );
      }
      else
      {
        labelInsertionPoint = new Point3d(
          textAlignmentReferencePoint.X - 14.1197 * 0.25 / scale,
          textAlignmentReferencePoint.Y,
          0
        );
      }
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        try
        {
          BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
          ObjectId locatorBlockId = bt["EQUIPMENT LOCATOR"];
          using (BlockReference acBlkRef = new BlockReference(firstClickPoint, locatorBlockId))
          {
            BlockTableRecord acCurSpaceBlkTblRec;
            acCurSpaceBlkTblRec =
              tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
            acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
            DynamicBlockReferencePropertyCollection pc =
              acBlkRef.DynamicBlockReferencePropertyCollection;
            foreach (DynamicBlockReferenceProperty prop in pc)
            {
              if (prop.PropertyName == "gmep_equip_id")
              {
                prop.Value = Id;
              }
              if (prop.PropertyName == "gmep_equip_parent_id" && ParentId != null)
              {
                prop.Value = ParentId;
              }
              if (prop.PropertyName == "gmep_equip_no")
              {
                prop.Value = Name;
              }
            }
            acBlkRef.Layer = "E-TXT1";
            tr.AddNewlyCreatedDBObject(acBlkRef, true);
          }
          tr.Commit();
        }
        catch (Autodesk.AutoCAD.Runtime.Exception ex)
        {
          tr.Commit();
        }
      }
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        try
        {
          BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
          ObjectId locatorBlockId = bt["EQUIP_MARKER"];
          using (BlockReference acBlkRef = new BlockReference(labelInsertionPoint, locatorBlockId))
          {
            BlockTableRecord acCurSpaceBlkTblRec;
            acCurSpaceBlkTblRec =
              tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
            acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
            DynamicBlockReferencePropertyCollection pc =
              acBlkRef.DynamicBlockReferencePropertyCollection;
            foreach (DynamicBlockReferenceProperty prop in pc)
            {
              if (prop.PropertyName == "gmep_equip_id")
              {
                prop.Value = Id;
              }
              if (prop.PropertyName == "gmep_equip_parent_id" && ParentId != null)
              {
                prop.Value = ParentId;
              }
            }
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
            attrDef.Position = labelInsertionPoint;
            attrDef.LockPositionInBlock = true;
            attrDef.Tag = Name;
            attrDef.IsMTextAttributeDefinition = false;
            attrDef.TextString = Name;
            attrDef.Justify = AttachmentPoint.MiddleCenter;
            attrDef.Visible = true;
            attrDef.Invisible = false;
            attrDef.Constant = false;
            attrDef.Height = 4.5 * 0.25 / scale;
            attrDef.WidthFactor = 0.85;
            attrDef.TextStyleId = gmepTextStyleId;
            attrDef.Layer = "0";

            AttributeReference attrRef = new AttributeReference();
            attrRef.SetAttributeFromBlock(attrDef, acBlkRef.BlockTransform);
            acBlkRef.AttributeCollection.AppendAttribute(attrRef);
            acBlkRef.Layer = "E-TXT1";
            acBlkRef.ScaleFactors = new Scale3d(0.25 / scale);
            tr.AddNewlyCreatedDBObject(acBlkRef, true);
          }
          tr.Commit();
        }
        catch (Autodesk.AutoCAD.Runtime.Exception ex)
        {
          tr.Commit();
        }
      }
      return firstClickPoint;
    }
  }

  public class LightingControl : PlaceableElectricalEntity
  {
    public string ControlType;
    public bool HasOccupancy;

    public LightingControl(string Id, string Name, string ControlTypeId, bool HasOccupancy)
    {
      this.Id = Id;
      this.Name = Name;
      this.ControlType = ControlTypeId;
      this.HasOccupancy = HasOccupancy;
    }
  }

  public class LightingFixture : PlaceableElectricalEntity
  {
    public int Voltage,
      Qty;
    public double Wattage,
      PaperSpaceScale,
      LabelTransformHX,
      LabelTransformHY,
      LabelTransformVX,
      LabelTransformVY;
    public string ControlId,
      Description,
      Mounting,
      Manufacturer,
      ModelNo,
      Notes;
    public bool EmCapable;

    public LightingFixture(
      string Id,
      string ParentId,
      string ParentName,
      string Name,
      string ControlId,
      string BlockName,
      int Voltage,
      double Wattage,
      string Description,
      int Qty,
      string Mounting,
      string Manufacturer,
      string ModelNo,
      string Notes,
      bool Rotate,
      double PaperSpaceScale,
      bool EmCapable,
      double LabelTransformHX,
      double LabelTransformHY,
      double LabelTransformVX,
      double LabelTransformVY
    )
    {
      this.Id = Id;
      this.ParentId = ParentId;
      this.ParentName = ParentName;
      this.Name = Name;
      this.BlockName = BlockName;
      this.Voltage = Voltage;
      this.Wattage = Wattage;
      this.ControlId = ControlId;
      this.Description = Description;
      this.Qty = Qty;
      this.Mounting = Mounting;
      this.Manufacturer = Manufacturer;
      this.ModelNo = ModelNo;
      this.Notes = Notes;
      this.Rotate = Rotate;
      this.PaperSpaceScale = PaperSpaceScale;
      this.EmCapable = EmCapable;
      this.LabelTransformHX = LabelTransformHX;
      this.LabelTransformHY = LabelTransformHY;
      this.LabelTransformVX = LabelTransformVX;
      this.LabelTransformVY = LabelTransformVY;
    }
  }

  public class DistributionBus : PlaceableElectricalEntity
  {
    public int AmpRating;

    public DistributionBus(
      string Id,
      string NodeId,
      string Status,
      int AmpRating,
      double AicRating,
      System.Drawing.Point NodePosition
    )
    {
      this.Id = Id;
      this.NodeId = NodeId;
      this.Status = Status;
      this.AmpRating = AmpRating;
      this.AicRating = AicRating;
      Name = $"{AmpRating}A Distrib. Bus";
      NodeType = NodeType.DistributionBus;
      this.NodePosition = NodePosition;
      BlockName = "GMEP DISTRIBUTION SECTION";
      Rotate = true;
    }
  }

  public class Service : PlaceableElectricalEntity
  {
    public int AmpRating;
    public string Voltage;

    public Service(
      string Id,
      string NodeId,
      string Status,
      int AmpRating,
      string Voltage,
      double AicRating,
      System.Drawing.Point NodePosition
    )
    {
      this.Id = Id;
      this.NodeId = NodeId;
      Name = $"{AmpRating}A {Voltage.Replace(" ", "V-")} Service";
      this.Status = Status;
      this.AmpRating = AmpRating;
      this.Voltage = Voltage;
      this.AicRating = AicRating;
      NodeType = NodeType.Service;
      this.NodePosition = NodePosition;
      BlockName = "GMEP SERVICE SECTION";
      Rotate = true;
    }
  }

  public class Equipment : PlaceableElectricalEntity
  {
    public string Description,
      Hp,
      Category;
    public int Voltage,
      MountingHeight,
      Circuit;
    public double Fla,
      Mca;
    public bool Is3Phase,
      HasPlug;

    public Equipment(
      string Id,
      string ParentId,
      string ParentName,
      string Name,
      string Description,
      string Category,
      int Voltage,
      double Fla,
      bool Is3Phase,
      int ParentDistance,
      double LocationX,
      double LocationY,
      float Mca,
      string Hp,
      int MountingHeight,
      int Circuit,
      bool HasPlug,
      bool Hidden
    )
    {
      this.Id = Id;
      this.ParentId = ParentId;
      this.ParentName = ParentName;
      this.Name = Name.ToUpper();
      this.Description = Description;
      this.Category = Category;
      this.Voltage = Voltage;
      this.Fla = Fla;
      this.Is3Phase = Is3Phase;
      this.Location = new Point3d(LocationX, LocationY, 0);
      this.ParentDistance = ParentDistance;
      this.Mca = Mca;
      this.Hp = Hp;
      this.MountingHeight = MountingHeight;
      this.Circuit = Circuit;
      this.HasPlug = HasPlug;
      this.IsHidden = Hidden;
    }
  }

  public class Disconnect : PlaceableElectricalEntity
  {
    public int AsSize;
    public int AfSize;
    public int NumPoles;

    public Disconnect(
      string Id,
      string ParentId,
      int ParentDistance,
      string NodeId,
      string Status,
      int AsSize,
      int AfSize,
      int NumPoles,
      double AicRating,
      System.Drawing.Point NodePosition
    )
    {
      this.Id = Id;
      this.ParentId = ParentId;
      this.ParentDistance = ParentDistance;
      this.NodeId = NodeId;
      this.AsSize = AsSize;
      this.AfSize = AfSize;
      this.NumPoles = NumPoles;
      this.Status = Status;
      this.AicRating = AicRating;
      Name = $"{AsSize}AS/{AfSize}AF/{NumPoles}P Disconnect";
      NodeType = NodeType.Disconnect;
      this.NodePosition = NodePosition;
      BlockName = "GMEP DISCONNECT";
      Rotate = false;
    }
  }

  public class Panel : PlaceableElectricalEntity
  {
    public int BusAmpRating;
    public int MainAmpRating;
    public string Voltage;
    public bool IsMlo;
    public List<PanelBreaker> Breakers;

    public Panel(
      string Id,
      string ParentId,
      string Name,
      int ParentDistance,
      double LocationX,
      double LocationY,
      int BusAmpRating,
      int MainAmpRating,
      bool IsMlo,
      string Voltage,
      double AicRating,
      bool IsHidden,
      string NodeId,
      string Status,
      System.Drawing.Point NodePosition
    )
    {
      this.Id = Id;
      this.ParentId = ParentId;
      this.Name = Name;
      this.ParentDistance = ParentDistance;
      Location = new Point3d(LocationX, LocationY, 0);
      this.BusAmpRating = BusAmpRating;
      this.MainAmpRating = MainAmpRating;
      this.IsMlo = IsMlo;
      this.Voltage = Voltage;
      this.AicRating = AicRating;
      this.IsHidden = IsHidden;
      this.NodeId = NodeId;
      this.Status = Status;
      NodeType = NodeType.Panel;
      this.NodePosition = NodePosition;
      BlockName = $"A$C26441056";
      Rotate = true;
    }
  }

  public class Transformer : PlaceableElectricalEntity
  {
    public double Kva;
    public string Voltage;

    public Transformer(
      string Id,
      string ParentId,
      string Name,
      int ParentDistance,
      double LocationX,
      double LocationY,
      double Kva,
      string Voltage,
      double AicRating,
      bool IsHidden,
      string NodeId,
      string Status,
      System.Drawing.Point NodePosition
    )
    {
      this.Id = Id;
      this.ParentId = ParentId;
      this.Name = Name;
      this.ParentDistance = ParentDistance;
      Location = new Point3d(LocationX, LocationY, 0);
      this.Kva = Kva;
      this.Voltage = Voltage;
      this.AicRating = AicRating;
      this.IsHidden = IsHidden;
      this.NodeId = NodeId;
      this.Status = Status;
      NodeType = NodeType.Transformer;
      this.NodePosition = NodePosition;
      BlockName = "GMEP TRANSFORMER";
      Rotate = false;
    }
  }
}
