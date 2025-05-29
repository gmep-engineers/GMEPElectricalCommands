using System;
using System.IO;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands.Equipment
{
  public class EquipmentCommands
  {
    private EquipmentDialogWindow EquipWindow;

    [CommandMethod("EQUIP")]
    public void EQUIP()
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Editor ed = doc.Editor;
      string fileName = Path.GetFileName(doc.Name).Substring(0, 6);
      if (!Regex.IsMatch(fileName, @"[0-9]{2}-[0-9]{3}"))
      {
        ed.WriteMessage("\nFilename invalid format. Filename must begin with GMEP project number.");
        return;
      }
      try
      {
        if (this.EquipWindow != null && !this.EquipWindow.IsDisposed)
        {
          this.EquipWindow.BringToFront();
        }
        else
        {
          this.EquipWindow = new EquipmentDialogWindow(this);
          this.EquipWindow.InitializeModal();
          this.EquipWindow.Show();
        }
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage("Error: " + ex.ToString());
      }
    }

    [CommandMethod("CloneEquip")]
    public void CloneEquip()
    {
      double scale = 12;
      if (
        CADObjectCommands.Scale <= 0
        && (CADObjectCommands.IsInModel() || CADObjectCommands.IsInLayoutViewport())
      )
      {
        CADObjectCommands.SetScale();
        if (CADObjectCommands.Scale <= 0)
          return;
      }
      if (CADObjectCommands.IsInModel() || CADObjectCommands.IsInLayoutViewport())
      {
        scale = CADObjectCommands.Scale;
      }
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Editor ed = doc.Editor;
      Database db = doc.Database;
      PromptPointOptions promptOptions = new PromptPointOptions("\nSelect source base point: ");
      PromptPointResult basePointPromptResult = ed.GetPoint(promptOptions);
      if (basePointPromptResult.Status != PromptStatus.OK)
        return;
      Point3d basePoint = basePointPromptResult.Value;

      PromptPointOptions clonePromptOptions = new PromptPointOptions("\nSelect clone base point: ");
      PromptPointResult clonePointPromptResult = ed.GetPoint(clonePromptOptions);
      if (clonePointPromptResult.Status != PromptStatus.OK)
        return;
      Point3d clonePoint = clonePointPromptResult.Value;

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
            BlockReference br = (BlockReference)tr.GetObject(id, OpenMode.ForWrite);

            if (
              br != null
              && br.IsDynamicBlock
              && br.DynamicBlockReferencePropertyCollection.Count > 0
            )
            {
              DynamicBlockReferencePropertyCollection pc =
                br.DynamicBlockReferencePropertyCollection;
              bool isLocator = false;
              string equipId = "";
              string parentId = "";
              string equipNo = "";
              foreach (DynamicBlockReferenceProperty prop in pc)
              {
                if (prop.PropertyName == "gmep_equip_locator" && prop.Value as string == "true")
                {
                  isLocator = true;
                }
                if (prop.PropertyName == "gmep_equip_id")
                {
                  equipId = prop.Value as string;
                }
                if (prop.PropertyName == "gmep_equip_parent_id")
                {
                  parentId = prop.Value as string;
                }
                if (prop.PropertyName == "gmep_equip_no")
                {
                  equipNo = prop.Value as string;
                }
              }
              if (isLocator)
              {
                BlockTableRecord curSpace = (BlockTableRecord)
                  tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                BlockTableRecord block = (BlockTableRecord)
                  tr.GetObject(bt[br.Name], OpenMode.ForRead);
                Point3d clonedBlockPoint = new Point3d(
                  br.Position.X + clonePoint.X - basePoint.X,
                  br.Position.Y + clonePoint.Y - basePoint.Y,
                  0
                );
                using (BlockReference cbr = new BlockReference(clonedBlockPoint, block.ObjectId))
                {
                  curSpace.AppendEntity(cbr);
                  tr.AddNewlyCreatedDBObject(cbr, true);

                  cbr.Layer = br.Layer;
                  cbr.Rotation = br.Rotation;

                  DynamicBlockReferencePropertyCollection cpc =
                    cbr.DynamicBlockReferencePropertyCollection;

                  foreach (DynamicBlockReferenceProperty prop in cpc)
                  {
                    if (prop.PropertyName == "gmep_equip_id")
                    {
                      prop.Value = equipId;
                    }
                    if (prop.PropertyName == "gmep_equip_parent_id")
                    {
                      prop.Value = parentId;
                    }
                    if (prop.PropertyName == "gmep_equip_no")
                    {
                      prop.Value = equipNo;
                    }
                    if (prop.PropertyName == "gmep_equip_locator")
                    {
                      prop.Value = "false";
                    }
                  }
                }
              }
              else if (!String.IsNullOrEmpty(equipId) && equipId != "0")
              {
                ObjectId locatorBlockId = bt["EQUIP_MARKER"];
                Point3d labelPoint = new Point3d(
                  br.Position.X + clonePoint.X - basePoint.X,
                  br.Position.Y + clonePoint.Y - basePoint.Y,
                  0
                );
                using (BlockReference acBlkRef = new BlockReference(labelPoint, locatorBlockId))
                {
                  BlockTableRecord acCurSpaceBlkTblRec;
                  acCurSpaceBlkTblRec =
                    tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                  acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                  DynamicBlockReferencePropertyCollection cpc =
                    acBlkRef.DynamicBlockReferencePropertyCollection;
                  foreach (DynamicBlockReferenceProperty prop in pc)
                  {
                    if (prop.PropertyName == "gmep_equip_id")
                    {
                      prop.Value = equipId;
                    }
                    if (prop.PropertyName == "gmep_equip_parent_id")
                    {
                      prop.Value = parentId;
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
                  string textString = "";
                  foreach (ObjectId objId in br.AttributeCollection)
                  {
                    AttributeReference ar =
                      tr.GetObject(objId, OpenMode.ForRead) as AttributeReference;

                    if (ar.Tag == "TAG")
                    {
                      textString = ar.TextString;
                    }
                  }

                  var textStyle = (TextStyleTableRecord)
                    tr.GetObject(gmepTextStyleId, OpenMode.ForRead);
                  double widthFactor = 1;
                  if (textStyle.FileName.ToLower().Contains("architxt"))
                  {
                    widthFactor = 0.85;
                  }

                  AttributeDefinition attrDef = new AttributeDefinition();
                  attrDef.Position = labelPoint;
                  attrDef.LockPositionInBlock = true;
                  attrDef.Tag = "TAG";
                  attrDef.IsMTextAttributeDefinition = false;
                  attrDef.TextString = textString;
                  attrDef.Justify = AttachmentPoint.MiddleCenter;
                  attrDef.Visible = true;
                  attrDef.Invisible = false;
                  attrDef.Constant = false;
                  attrDef.Height = 4.5 * 0.25 / scale;
                  attrDef.WidthFactor = widthFactor;
                  attrDef.TextStyleId = gmepTextStyleId;
                  attrDef.Layer = "0";

                  AttributeReference attrRef = new AttributeReference();
                  attrRef.SetAttributeFromBlock(attrDef, acBlkRef.BlockTransform);
                  acBlkRef.AttributeCollection.AppendAttribute(attrRef);
                  acBlkRef.Layer = "E-TXT1";
                  acBlkRef.ScaleFactors = new Scale3d(0.25 / scale);
                  tr.AddNewlyCreatedDBObject(acBlkRef, true);
                }
              }
            }
          }
          catch { }
          try
          {
            Line line = (Line)tr.GetObject(id, OpenMode.ForRead);
            if (line != null)
            {
              ObjectId fieldId = line.GetField("gmep_equip_id");
              ObjectId isCloneId = line.GetField("is_clone");
              if (fieldId != null && isCloneId != null)
              {
                Field field = (Field)tr.GetObject(fieldId, OpenMode.ForRead);
                Field isClone = (Field)tr.GetObject(isCloneId, OpenMode.ForRead);
                if (isClone != null && isClone.GetFieldCode() == "false")
                {
                  string equipId = field.GetFieldCode();
                  Point3d startPoint = new Point3d(
                    line.StartPoint.X + clonePoint.X - basePoint.X,
                    line.StartPoint.Y + clonePoint.Y - basePoint.Y,
                    0
                  );
                  Point3d endPoint = new Point3d(
                    line.EndPoint.X + clonePoint.X - basePoint.X,
                    line.EndPoint.Y + clonePoint.Y - basePoint.Y,
                    0
                  );
                  Line clonedLine = new Line(startPoint, endPoint);
                  Field newField = new Field(equipId);
                  Field newIsClone = new Field("true");
                  clonedLine.SetField("gmep_equip_id", newField);
                  clonedLine.SetField("is_clone", newIsClone);
                  clonedLine.Layer = line.Layer;
                  btr.AppendEntity(clonedLine);
                  tr.AddNewlyCreatedDBObject(clonedLine, true);
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
              if (
                text.Hyperlinks.Count > 0
                && text.Hyperlinks[0].SubLocation.Contains("gmep_equip")
                && !text.Hyperlinks[0].SubLocation.Contains("cloned")
              )
              {
                ObjectId clonedTextId = GeneralCommands.CreateAndPositionText(
                  tr,
                  text.TextString,
                  "gmep",
                  0.0938 * 12 / CADObjectCommands.Scale,
                  1,
                  2,
                  "E-TXT1",
                  new Point3d(
                    text.AlignmentPoint.X + clonePoint.X - basePoint.X,
                    text.AlignmentPoint.Y + clonePoint.Y - basePoint.Y,
                    0
                  ),
                  text.HorizontalMode,
                  text.VerticalMode,
                  text.Justify
                );
                var clonedTextObj = (DBText)tr.GetObject(clonedTextId, OpenMode.ForWrite);
                HyperLink hyperLink = new HyperLink();
                hyperLink.SubLocation = text.Hyperlinks[0].SubLocation + "cloned";
                clonedTextObj.Hyperlinks.Add(hyperLink);
              }
            }
          }
          catch { }
        }
        tr.Commit();
      }
    }
  }
}
