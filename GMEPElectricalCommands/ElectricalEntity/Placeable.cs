using System;
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
    public string TableName;
    public ObjectId BlockId;
    public int AmpRating;
    public string Voltage;
    public double LoadAmperage;
    public double Kva;

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
          try
          {
            Line line = (Line)tr.GetObject(id, OpenMode.ForWrite);
            if (line != null)
            {
              ObjectId fieldId = line.GetField("gmep_equip_id");
              if (fieldId != null)
              {
                Field field = (Field)tr.GetObject(fieldId, OpenMode.ForWrite);
                if (field != null && field.GetFieldCode() == Id)
                {
                  line.Erase();
                }
              }
            }
          }
          catch { }
          try
          {
            DBText text = (DBText)tr.GetObject(id, OpenMode.ForWrite);
            if (text != null)
            {
              if (text.Hyperlinks.Count > 0 && text.Hyperlinks[0].SubLocation == Id)
              {
                text.Erase();
              }
            }
          }
          catch { }
        }
        tr.Commit();
      }
      double scale = 12;
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
        BlockJig blockJig = new BlockJig(Name);

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

          using (BlockReference br = new BlockReference(firstClickPoint, block.ObjectId))
          {
            if (blockPromptResult.Status == PromptStatus.OK)
            {
              BlockTableRecord currentSpace =
                tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
              br.ScaleFactors = new Scale3d(0.25 / scaleFactor);
              br.Layer = IsExisting() ? "E-SYM-EXISTING" : "E-SYM1";
            }

            if (Rotate)
            {
              RotateJig rotateJig = new RotateJig(br);
              PromptResult rotatePromptResult = ed.Drag(rotateJig);

              if (rotatePromptResult.Status != PromptStatus.OK)
              {
                return null;
              }
              rotation = br.Rotation;
            }
            curSpace.AppendEntity(br);
            tr.AddNewlyCreatedDBObject(br, true);
            BlockId = br.Id;
            Location = firstClickPoint;
            tr.Commit();
          }
        }
        else
        {
          return null;
        }
      }

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        try
        {
          BlockReference br = (BlockReference)tr.GetObject(BlockId, OpenMode.ForWrite);
          if (
            br != null
            && br.IsDynamicBlock
            && br.DynamicBlockReferencePropertyCollection.Count > 0
          )
          {
            DynamicBlockReferencePropertyCollection pc = br.DynamicBlockReferencePropertyCollection;
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
              if (prop.PropertyName == "gmep_equip_locator")
              {
                prop.Value = "true";
              }
            }
            tr.Commit();
          }
        }
        catch { }
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
      if (NodeType == NodeType.Transformer)
      {
        firstClickPoint = new Point3d(firstClickPoint.X, firstClickPoint.Y + 19.9429, 0);
      }
      if (NodeType == NodeType.Disconnect)
      {
        firstClickPoint = new Point3d(firstClickPoint.X, firstClickPoint.Y + 2.5 * 0.25 / scale, 0);
      }
      LabelJig jig = new LabelJig(firstClickPoint, Id);
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
        DynamicLineJig lineJig = new DynamicLineJig(jig.endPoint, scale, Id);
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

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        try
        {
          if (NodeType == NodeType.Panel)
          {
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
            //ObjectId textId = GeneralCommands.CreateAndPositionText(
            //  tr,
            //  GetStatusAbbr(),
            //  "gmep",
            //  0.0938 * 12 / scale,
            //  0.85,
            //  2,
            //  "E-TXT1",
            //  new Point3d(
            //    labelInsertionPoint.X - (0.85 / scale),
            //    labelInsertionPoint.Y + (1.25 / scale),
            //    0
            //  ),
            //  TextHorizontalMode.TextCenter,
            //  TextVerticalMode.TextBase,
            //  AttachmentPoint.BaseLeft
            //);
            ObjectId statusAndAmpTextId = GeneralCommands.CreateAndPositionText(
              tr,
              GetStatusAbbr() + AmpRating.ToString() + "A",
              "gmep",
              0.0938 * 12 / scale,
              0.85,
              2,
              "E-TXT1",
              new Point3d(labelInsertionPoint.X, labelInsertionPoint.Y - (1.5 * 1.5 / scale), 0),
              TextHorizontalMode.TextCenter,
              TextVerticalMode.TextBase,
              AttachmentPoint.BaseCenter
            );
            string voltageString = Voltage.Substring(0, 7) + "V";
            ObjectId voltageTextId = GeneralCommands.CreateAndPositionText(
              tr,
              voltageString,
              "gmep",
              0.0938 * 12 / scale,
              0.85,
              2,
              "E-TXT1",
              new Point3d(labelInsertionPoint.X, labelInsertionPoint.Y - (2.5 * 1.5 / scale), 0),
              TextHorizontalMode.TextCenter,
              TextVerticalMode.TextBase,
              AttachmentPoint.BaseCenter
            );
            string phaseWireString = Phase == 3 ? "3\u0081-4W" : "1\u0081-3W";
            ObjectId phaseWireTextId = GeneralCommands.CreateAndPositionText(
              tr,
              phaseWireString,
              "gmep",
              0.0938 * 12 / scale,
              0.85,
              2,
              "E-TXT1",
              new Point3d(labelInsertionPoint.X, labelInsertionPoint.Y - (3.5 * 1.5 / scale), 0),
              TextHorizontalMode.TextCenter,
              TextVerticalMode.TextBase,
              AttachmentPoint.BaseCenter
            );

            // set hyperlink
            var statusAndAmpTextObj = (DBText)tr.GetObject(statusAndAmpTextId, OpenMode.ForWrite);
            var voltageTextObj = (DBText)tr.GetObject(voltageTextId, OpenMode.ForWrite);
            var phaseWireTextObj = (DBText)tr.GetObject(phaseWireTextId, OpenMode.ForWrite);
            // this is the quickest way to add a custom attribute to DBText without
            // having to do a bunch of bloated AutoCAD database nonsense
            HyperLink customAttr = new HyperLink();
            customAttr.SubLocation = Id;
            statusAndAmpTextObj.Hyperlinks.Add(customAttr);
            voltageTextObj.Hyperlinks.Add(customAttr);
            phaseWireTextObj.Hyperlinks.Add(customAttr);

            BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            ObjectId locatorBlockId = bt["EQUIP_MARKER"];
            using (
              BlockReference acBlkRef = new BlockReference(labelInsertionPoint, locatorBlockId)
            )
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
          else
          {
            // create text above line
            Point3d labelInsertionPoint;
            double verticalAdjustment = -0.5 / scale;
            labelInsertionPoint = new Point3d(
              textAlignmentReferencePoint.X,
              textAlignmentReferencePoint.Y + verticalAdjustment,
              0
            );
            ObjectId textId;
            if (textAlignmentReferencePoint.X > comparisonPoint.X)
            {
              textId = GeneralCommands.CreateAndPositionText(
                tr,
                GetStatusAbbr() + Name.ToUpper(),
                "gmep",
                0.0938 * 12 / scale,
                0.85,
                2,
                "E-TXT1",
                new Point3d(labelInsertionPoint.X + (0.5 / scale), labelInsertionPoint.Y, 0),
                TextHorizontalMode.TextCenter,
                TextVerticalMode.TextBase,
                AttachmentPoint.BaseLeft
              );
            }
            else
            {
              textId = GeneralCommands.CreateAndPositionText(
                tr,
                GetStatusAbbr() + Name.ToUpper(),
                "gmep",
                0.0938 * 12 / scale,
                0.85,
                2,
                "E-TXT1",
                new Point3d(labelInsertionPoint.X - (0.5 / scale), labelInsertionPoint.Y, 0),
                TextHorizontalMode.TextCenter,
                TextVerticalMode.TextBase,
                AttachmentPoint.BaseRight
              );
            }
            var text = (DBText)tr.GetObject(textId, OpenMode.ForWrite);
            // this is the quickest way to add a custom attribute to DBText without
            // having to do a bunch of bloated AutoCAD database nonsense
            HyperLink customAttr = new HyperLink();
            customAttr.SubLocation = Id;
            text.Hyperlinks.Add(customAttr);
            tr.Commit();
          }
        }
        catch (Autodesk.AutoCAD.Runtime.Exception ex)
        {
          tr.Commit();
        }
      }
      return firstClickPoint;
    }

    public void SetBlockId(List<ObjectId> dynamicBlockIds)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;

      Database db = doc.Database;
      Editor ed = doc.Editor;
      Transaction tr = db.TransactionManager.StartTransaction();
      using (tr)
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        foreach (ObjectId id in dynamicBlockIds)
        {
          try
          {
            BlockReference br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
            if (br != null && br.IsDynamicBlock)
            {
              DynamicBlockReferencePropertyCollection pc =
                br.DynamicBlockReferencePropertyCollection;
              bool isEquip = false;
              ObjectId blockId = new ObjectId();
              foreach (DynamicBlockReferenceProperty prop in pc)
              {
                if (prop.PropertyName == "gmep_equip_id" && prop.Value as string == Id)
                {
                  blockId = id;
                }
                if (prop.PropertyName == "gmep_equip_locator" && prop.Value as string == "true")
                {
                  isEquip = true;
                }
              }
              if (isEquip)
              {
                BlockId = blockId;
              }
            }
          }
          catch { }
        }
      }
    }

    private bool ResetLocation()
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;

      Database db = doc.Database;
      Editor ed = doc.Editor;
      Transaction tr = db.TransactionManager.StartTransaction();
      using (tr)
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        var modelSpace = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

        foreach (ObjectId id in modelSpace)
        {
          try
          {
            BlockReference br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
            if (br != null && br.IsDynamicBlock)
            {
              DynamicBlockReferencePropertyCollection pc =
                br.DynamicBlockReferencePropertyCollection;
              foreach (DynamicBlockReferenceProperty prop in pc)
              {
                if (prop.PropertyName == "gmep_equip_id" && prop.Value as string != "0") { }
              }
            }
          }
          catch { }
        }

        BlockReference _br = (BlockReference)tr.GetObject(BlockId, OpenMode.ForRead);
        if (Location != _br.Position)
        {
          Location = _br.Position;
          return true;
        }
      }
      return false;
    }
  }

  public class LightingControl : PlaceableElectricalEntity
  {
    public string ControlType;
    public bool HasOccupancy;

    public LightingControl(string Id, string Name, string ControlType, bool HasOccupancy)
    {
      this.Id = Id;
      this.Name = Name;
      this.ControlType = ControlType;
      this.HasOccupancy = HasOccupancy;
    }
  }

  public class LightingFixture : PlaceableElectricalEntity
  {
    public int Voltage,
      Qty,
      Circuit,
      Pole;
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
      double LabelTransformVY,
      int Circuit
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
      this.Circuit = Circuit;
      this.Pole = 1;
    }
  }

  public class DistributionBus : PlaceableElectricalEntity
  {
    public DistributionBus(
      string Id,
      string NodeId,
      string Status,
      int AmpRating,
      double AicRating,
      System.Drawing.Point NodePosition,
      Point3d Location
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
      TableName = "electrical_distribution_buses";
      Rotate = true;
      this.Location = Location;
    }
  }

  public class Service : PlaceableElectricalEntity
  {
    public Service(
      string Id,
      string NodeId,
      string Status,
      int AmpRating,
      string Voltage,
      double AicRating,
      System.Drawing.Point NodePosition,
      Point3d Location
    )
    {
      this.Id = Id;
      this.NodeId = NodeId;
      Name =
        $"{AmpRating}A {Voltage.Replace(" ", "V-")}"
        + "\u0081"
        + $"-{(Voltage.Contains("3") ? "4W" : "3W")} Service";
      this.Status = Status;
      this.AmpRating = AmpRating;
      this.Voltage = Voltage;
      this.AicRating = AicRating;
      NodeType = NodeType.Service;
      this.NodePosition = NodePosition;
      BlockName = "GMEP SERVICE SECTION";
      TableName = "electrical_services";
      Rotate = true;
      this.Location = Location;
      LineVoltage = 208;
      if (Voltage.Contains("480"))
      {
        LineVoltage = 480;
      }
      if (Voltage.Contains("240"))
      {
        LineVoltage = 240;
      }
      Phase = 1;
      if (Voltage.Contains("3"))
      {
        Phase = 3;
      }
    }
  }

  public class Equipment : PlaceableElectricalEntity
  {
    public string Description,
      Hp,
      Category;

    public int MountingHeight,
      Circuit,
      Pole;
    public double Fla,
      Mca;
    public bool Is3Phase,
      HasPlug;

    public Equipment(
      string Id,
      string NodeId,
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
      bool Hidden,
      string Status
    )
    {
      this.Id = Id;
      this.NodeId = NodeId;
      this.ParentId = ParentId;
      this.ParentName = ParentName;
      this.Name = Name.ToUpper();
      this.Description = Description;
      this.Category = Category;
      this.Voltage = Voltage.ToString();
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
      this.Status = Status;
      this.NodeType = NodeType.Equipment;
      TableName = "electrical_equipment";
      Pole = SetPole(Is3Phase, Voltage);
    }

    private int SetPole(bool is3Phase, int voltage)
    {
      int pole = 3;
      if (is3Phase == false)
      {
        if (voltage == 115 || voltage == 120 || voltage == 277)
        {
          pole = 1;
        }
        else
        {
          pole = 2;
        }
      }
      return pole;
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
      System.Drawing.Point NodePosition,
      Point3d Location
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
      AmpRating = AfSize;
      Name = $"{AsSize}AS/{AfSize}AF/{NumPoles}P Disconnect";
      NodeType = NodeType.Disconnect;
      this.NodePosition = NodePosition;
      BlockName = "GMEP DISCONNECT";
      Rotate = false;
      TableName = "electrical_disconnects";
      this.Location = Location;
    }
  }

  public class Panel : PlaceableElectricalEntity
  {
    public int BusAmpRating;
    public int MainAmpRating;
    public bool IsMlo;
    public bool IsRecessed;
    public int NumBreakers;
    public List<PanelBreaker> Breakers;
    public int Circuit;
    public int Pole;

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
      bool IsRecessed,
      string Voltage,
      double LoadAmperage,
      double Kva,
      double AicRating,
      bool IsHidden,
      string NodeId,
      string Status,
      System.Drawing.Point NodePosition,
      int NumBreakers,
      int Circuit
    )
    {
      this.Id = Id;
      this.ParentId = ParentId;
      this.Name = Name.ToUpper().Replace("PANEL", "").Trim();
      this.ParentDistance = ParentDistance;
      Location = new Point3d(LocationX, LocationY, 0);
      this.BusAmpRating = BusAmpRating;
      this.MainAmpRating = MainAmpRating;
      AmpRating =
        IsMlo ? BusAmpRating
        : MainAmpRating < BusAmpRating ? MainAmpRating
        : BusAmpRating;
      this.IsMlo = IsMlo;
      this.IsRecessed = IsRecessed;
      this.Voltage = Voltage;
      this.AicRating = AicRating;
      this.IsHidden = IsHidden;
      this.NodeId = NodeId;
      this.Status = Status;
      NodeType = NodeType.Panel;
      this.NodePosition = NodePosition;
      BlockName = IsRecessed ? "GMEP RECESSED PANEL" : "GMEP SURFACE PANEL";
      Rotate = true;
      TableName = "electrical_panels";
      this.NumBreakers = NumBreakers;
      this.Circuit = Circuit;
      this.Pole = SetPole(Voltage);
      this.LoadAmperage = LoadAmperage;
      this.Kva = Kva;
    }

    public int SetPole(string voltage)
    {
      if (voltage == "120/240 1" || voltage == "120/208 1")
      {
        return 2;
      }
      return 3;
    }
  }

  public class Transformer : PlaceableElectricalEntity
  {
    public double Kva;
    public string Voltage;
    public int Circuit;
    public int Pole;
    public double OutputLineVoltage;

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
      System.Drawing.Point NodePosition,
      int Circuit
    )
    {
      this.Id = Id;
      this.ParentId = ParentId;
      this.ParentDistance = ParentDistance;
      Location = new Point3d(LocationX, LocationY, 0);
      this.Kva = Kva;
      this.Voltage = Voltage + "\u0081" + $"-{(Voltage.Contains("3") ? "4W" : "3W")}";
      this.Name = $"{Kva}KVA, {this.Voltage} XFMR '{Name.ToUpper()}'";
      this.AicRating = AicRating;
      this.IsHidden = IsHidden;
      this.NodeId = NodeId;
      this.Status = Status;
      NodeType = NodeType.Transformer;
      this.NodePosition = NodePosition;
      BlockName = "GMEP TRANSFORMER";
      Rotate = false;
      TableName = "electrical_transformers";
      this.Circuit = Circuit;
      this.Pole = SetPole(Voltage);
      LineVoltage = Double.Parse(Voltage.Split('-')[0].Replace("V", ""));
      OutputLineVoltage = Double.Parse(
        Voltage.Split('-')[1].Replace("120/", "").Replace("277/", "").Replace("V", "")
      );
      Phase = Int32.Parse(Voltage.Split('-')[2]);
    }

    public int SetPole(string voltage)
    {
      if (voltage == "240V-120/208V-1" || voltage == "208V-120/240V-1")
      {
        return 2;
      }
      return 3;
    }
  }
}
