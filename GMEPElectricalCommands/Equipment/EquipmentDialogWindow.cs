using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Dreambuild.AutoCAD;
using GMEPElectricalCommands.GmepDatabase;

namespace ElectricalCommands.Equipment
{
  public struct Equipment
  {
    public string equipId,
      parentId,
      parentName,
      objectId;
    public string equipNo,
      description,
      category;
    public int voltage;
    public bool is3Phase;
    public int parentDistance;
    public Point3d loc;

    public Equipment(
      string eqId = "",
      string pId = "",
      string pName = "",
      string eqName = "",
      string desc = "",
      string cat = "",
      int volts = 0,
      bool is3Ph = false,
      int pDist = 0,
      double xLoc = 0,
      double yLoc = 0
    )
    {
      equipId = eqId;
      parentId = pId;
      parentName = pName;
      equipNo = eqName;
      description = desc;
      category = cat;
      voltage = volts;
      is3Phase = is3Ph;
      loc = new Point3d(xLoc, yLoc, 0);
      parentDistance = pDist;
    }
  }

  public struct Feeder
  {
    public string feederId,
      name,
      parentId,
      type;
    public Point3d loc;
    public int parentDistance;

    public Feeder(
      string id,
      string pId,
      string n,
      string t,
      int fdrDist = 0,
      double xLoc = 0,
      double yLoc = 0
    )
    {
      feederId = id;
      parentId = pId;
      name = n;
      type = t;
      parentDistance = fdrDist;
      loc = new Point3d(xLoc, yLoc, 0);
    }
  }

  public partial class EquipmentDialogWindow : Form
  {
    private string filterPanel;
    private string filterVoltage;
    private string filterPhase;
    private string filterEquipNo;
    private string filterCategory;
    private List<ListViewItem> equipmentListViewList;
    private List<Equipment> equipmentList;
    private List<ListViewItem> feederListViewList;
    private List<Feeder> feederList;
    private string projectId;

    public GmepDatabase db = new GmepDatabase();

    public EquipmentDialogWindow(EquipmentCommands EquipCommands)
    {
      InitializeComponent();
    }

    public void InitializeModal()
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Core
        .Application
        .DocumentManager
        .MdiActiveDocument;
      string fileName = Path.GetFileName(doc.Name);
      //string projectNo = Regex.Match(fileName, @"[0-9]{2}-[0-9]{3}").Value;
      string projectNo = "24-123";
      projectId = db.GetProjectId(projectNo);
      feederList = db.GetFeeders(projectId);
      equipmentList = db.GetEquipment(projectId);
      CreateEquipmentListView();
      CreateFeederListView();
    }

    private void CreateEquipmentListView()
    {
      equipmentListView.View = View.Details;
      equipmentListView.FullRowSelect = true;
      foreach (Equipment equipment in equipmentList)
      {
        ListViewItem item = new ListViewItem(equipment.equipNo, 0);
        item.SubItems.Add(equipment.description);
        item.SubItems.Add(equipment.category);
        item.SubItems.Add(equipment.parentName);
        item.SubItems.Add(equipment.parentDistance.ToString());
        item.SubItems.Add(equipment.voltage.ToString());
        item.SubItems.Add(equipment.is3Phase ? "3" : "1");
        item.SubItems.Add(
          Math.Round(equipment.loc.X, 1).ToString()
            + ", "
            + Math.Round(equipment.loc.Y, 1).ToString()
        );
        item.SubItems.Add(equipment.equipId);
        item.SubItems.Add(equipment.parentId);
        equipmentListView.Items.Add(item);
      }
      equipmentListView.Columns.Add("Equip #", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Description", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Category", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Panel", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Panel Distance", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Voltage", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Phase", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Location", -2, HorizontalAlignment.Left);
    }

    private void CreateFeederListView()
    {
      feederListView.View = View.Details;
      feederListView.FullRowSelect = true;
      foreach (Feeder feeder in feederList)
      {
        ListViewItem item = new ListViewItem(feeder.name, 0);
        item.SubItems.Add(feeder.type);
        item.SubItems.Add(
          Math.Round(feeder.loc.X, 1).ToString() + ", " + Math.Round(feeder.loc.Y, 1).ToString()
        );
        feederListView.Items.Add(item);
      }
      feederListView.Columns.Add("Name", -2, HorizontalAlignment.Left);
      feederListView.Columns.Add("Type", -2, HorizontalAlignment.Left);
      feederListView.Columns.Add("Location", -2, HorizontalAlignment.Left);
    }

    private void FilterPanelComboBox_SelectedIndexChanged(object sender, EventArgs e)
    {
      filterPanel = GeneralCommands.GetComboBoxValue(filterPanelComboBox);
    }

    private void FilterVoltageComboBox_SelectedIndexChanged(object sender, EventArgs e)
    {
      filterVoltage = GeneralCommands.GetComboBoxValue(filterVoltageComboBox);
    }

    private Point3d PlaceEquipment(string equipId, string parentId, string equipNo)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      Point3d point;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
        var promptOptions = new PromptPointOptions("\nSelect point for " + equipNo + ": ");
        var promptResult = ed.GetPoint(promptOptions);
        if (promptResult.Status != PromptStatus.OK)
          return new Point3d();

        // Initial point
        point = promptResult.Value;
        try
        {
          ObjectId blockId = bt["EQUIP_MARKER"];
          using (BlockReference acBlkRef = new BlockReference(point, blockId))
          {
            BlockTableRecord acCurSpaceBlkTblRec;
            acCurSpaceBlkTblRec =
              tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
            acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
            //tr.AddNewlyCreatedDBObject(acBlkRef, true);

            DynamicBlockReferencePropertyCollection pc =
              acBlkRef.DynamicBlockReferencePropertyCollection;
            Equipment eq = new Equipment();
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

            AttributeDefinition attrDef = new AttributeDefinition();
            attrDef.Position = point;
            attrDef.LockPositionInBlock = true;
            attrDef.Tag = equipNo;
            attrDef.IsMTextAttributeDefinition = false;
            attrDef.TextString = equipNo;
            attrDef.Justify = AttachmentPoint.MiddleCenter;
            attrDef.Visible = true;
            attrDef.Invisible = false;
            attrDef.Constant = false;
            attrDef.Height = 4.5;
            attrDef.WidthFactor = 0.85;
            attrDef.Layer = "DEFPOINTS";

            acCurSpaceBlkTblRec.AppendEntity(attrDef);

            tr.AddNewlyCreatedDBObject(attrDef, true);

            AttributeReference attrRef = new AttributeReference();
            attrRef.TextString = attrDef.TextString;
            attrRef.Invisible = false;
            attrRef.Height = 4.5;
            attrRef.Visible = false;

            attrRef.SetAttributeFromBlock(attrDef, acBlkRef.BlockTransform);
            acBlkRef.AttributeCollection.AppendAttribute(attrRef);
            acBlkRef.Layer = "DEFPOINTS";
            tr.AddNewlyCreatedDBObject(acBlkRef, true);
          }
          tr.Commit();
        }
        catch (Autodesk.AutoCAD.Runtime.Exception ex)
        {
          tr.Commit();
        }
      }
      return point;
    }

    private void CalculateDistances()
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
              bool addEquip = false;
              Equipment eq = new Equipment();
              foreach (DynamicBlockReferenceProperty prop in pc)
              {
                if (prop.PropertyName == "gmep_equip_id" && prop.Value as string != "0")
                {
                  addEquip = true;
                  eq.equipId = prop.Value as string;
                }
                if (prop.PropertyName == "gmep_equip_parent_id" && prop.Value as string != "0")
                {
                  addEquip = true;
                  eq.parentId = prop.Value as string;
                }
              }
              eq.loc = br.Position;
              if (addEquip)
              {
                equipmentList.Add(eq);
              }
            }
          }
          catch { }
        }
      }
      for (int i = 0; i < equipmentList.Count; i++)
      {
        for (int j = 0; j < equipmentList.Count; j++)
        {
          if (equipmentList[j].parentId == equipmentList[i].equipId)
          {
            Equipment equip = equipmentList[j];
            equip.parentDistance = Convert.ToInt32(
              equipmentList[j].loc.DistanceTo(equipmentList[i].loc)
            );
          }
        }
      }
    }

    private void filterClearButton_Click(object sender, EventArgs e)
    {
      CalculateDistances();
    }

    private void EquipmentListView_MouseDoubleClick(object sender, MouseEventArgs e)
    {
      using (
        DocumentLock docLock =
          Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
      )
      {
        Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.WindowState = Autodesk
          .AutoCAD
          .Windows
          .Window
          .State
          .Maximized;
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Window.Focus();
        int numSubItems = equipmentListView.SelectedItems[0].SubItems.Count;
        Point3d p = PlaceEquipment(
          equipmentListView.SelectedItems[0].SubItems[numSubItems - 2].Text,
          equipmentListView.SelectedItems[0].SubItems[numSubItems - 1].Text,
          equipmentListView.SelectedItems[0].Text
        );
        for (int i = 0; i < equipmentList.Count; i++)
        {
          Equipment equipment = equipmentList[i];
          if (
            equipmentList[i].equipId
            == equipmentListView.SelectedItems[0].SubItems[numSubItems - 2].Text
          )
          {
            equipment.loc = p;
          }
        }
      }
    }

    private void PlaceSelectedButton_Click(object sender, EventArgs e)
    {
      using (
        DocumentLock docLock =
          Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
      )
      {
        Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.WindowState = Autodesk
          .AutoCAD
          .Windows
          .Window
          .State
          .Maximized;
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Window.Focus();
        int numSubItems = equipmentListView.SelectedItems[0].SubItems.Count;
        Dictionary<string, Point3d> equipLocs = new Dictionary<string, Point3d>();
        foreach (ListViewItem item in equipmentListView.SelectedItems)
        {
          Point3d p = PlaceEquipment(
            item.SubItems[numSubItems - 2].Text,
            item.SubItems[numSubItems - 1].Text,
            item.Text
          );
          equipLocs[item.SubItems[numSubItems - 2].Text] = p;
        }

        for (int i = 0; i < equipmentList.Count; i++)
        {
          Equipment equip = equipmentList[i];
          if (equipLocs.ContainsKey(equip.equipId))
          {
            equip.loc = equipLocs[equip.equipId];
          }
          equipmentList[i] = equip;
        }
        // update list views with new values
        // update values in sql
      }
    }
  }
}
