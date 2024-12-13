using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Input;
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
      string eqNo = "",
      string desc = "",
      string cat = "",
      int volts = 0,
      bool is3Ph = false,
      int pDist = -1,
      double xLoc = 0,
      double yLoc = 0
    )
    {
      equipId = eqId;
      parentId = pId;
      parentName = pName;
      equipNo = eqNo.ToUpper();
      description = desc;
      category = cat;
      voltage = volts;
      is3Phase = is3Ph;
      loc = new Point3d(xLoc, yLoc, 0);
      parentDistance = pDist;
    }
  }

  public struct Panel
  {
    public string equipId,
      name,
      parentId;
    public Point3d loc;
    public int parentDistance;

    public Panel(string id, string pId, string n, int pDist = -1, double xLoc = 0, double yLoc = 0)
    {
      equipId = id;
      parentId = pId;
      name = n;
      parentDistance = pDist;
      loc = new Point3d(xLoc, yLoc, 0);
    }
  }

  public struct PooledEquipment
  {
    public string equipId;
    public string parentId;
    public Point3d loc;
    public int parentDistance;

    public PooledEquipment(string eqId = "0", string pId = "0", Point3d p = new Point3d())
    {
      equipId = eqId;
      parentId = pId;
      loc = p;
      parentDistance = -1;
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
    private List<ListViewItem> panelListViewList;
    private List<Panel> panelList;
    private string projectId;
    private bool isLoading;

    public GmepDatabase gmepDb = new GmepDatabase();

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
      projectId = gmepDb.GetProjectId(projectNo);
      panelList = gmepDb.GetPanels(projectId);
      equipmentList = gmepDb.GetEquipment(projectId);
      CreateEquipmentListView();
      CreatePanelListView();
      ResetLocations();
      CalculateDistances();
      isLoading = false;
    }

    private void CreateEquipmentListView(bool updateOnly = false)
    {
      if (updateOnly)
      {
        equipmentListView.Items.Clear();
      }
      equipmentListView.View = View.Details;
      equipmentListView.FullRowSelect = true;
      foreach (Equipment equipment in equipmentList)
      {
        if (!String.IsNullOrEmpty(filterPanel) && equipment.parentName != filterPanel)
        {
          continue;
        }
        if (!String.IsNullOrEmpty(filterVoltage) && equipment.voltage.ToString() != filterVoltage)
        {
          continue;
        }
        if (
          !String.IsNullOrEmpty(filterPhase)
          && filterPhase.ToString() == "1"
          && equipment.is3Phase
        )
        {
          continue;
        }
        if (
          !String.IsNullOrEmpty(filterPhase)
          && filterPhase.ToString() == "3"
          && !equipment.is3Phase
        )
        {
          continue;
        }
        if (!String.IsNullOrEmpty(filterEquipNo) && equipment.equipNo != filterEquipNo)
        {
          continue;
        }
        if (
          !String.IsNullOrEmpty(filterCategory)
          && equipment.category.ToUpper() != filterCategory.ToUpper()
        )
        {
          continue;
        }
        ListViewItem item = new ListViewItem(equipment.equipNo, 0);
        item.SubItems.Add(equipment.description);
        item.SubItems.Add(equipment.category);
        item.SubItems.Add(equipment.parentName);
        if (equipment.parentDistance == -1)
        {
          item.SubItems.Add("Not Set");
        }
        else
        {
          item.SubItems.Add(equipment.parentDistance.ToString() + "'");
        }
        item.SubItems.Add(equipment.voltage.ToString());
        item.SubItems.Add(equipment.is3Phase ? "3" : "1");
        if (equipment.loc.X == 0 && equipment.loc.Y == 0)
        {
          item.SubItems.Add("Not Set");
        }
        else
        {
          item.SubItems.Add(
            Math.Round(equipment.loc.X / 12, 1).ToString()
              + ", "
              + Math.Round(equipment.loc.Y / 12, 1).ToString()
          );
        }
        item.SubItems.Add(equipment.equipId);
        item.SubItems.Add(equipment.parentId);
        equipmentListView.Items.Add(item);
      }
      if (!updateOnly)
      {
        equipmentListView.Columns.Add("Equip #", -2, HorizontalAlignment.Left);
        equipmentListView.Columns.Add("Description", -2, HorizontalAlignment.Left);
        equipmentListView.Columns.Add("Category", -2, HorizontalAlignment.Left);
        equipmentListView.Columns.Add("Panel", -2, HorizontalAlignment.Left);
        equipmentListView.Columns.Add("Panel Distance", -2, HorizontalAlignment.Left);
        equipmentListView.Columns.Add("Voltage", -2, HorizontalAlignment.Left);
        equipmentListView.Columns.Add("Phase", -2, HorizontalAlignment.Left);
        equipmentListView.Columns.Add("Location", -2, HorizontalAlignment.Left);
      }
    }

    private void CreatePanelListView(bool updateOnly = false)
    {
      if (updateOnly)
      {
        panelListView.Items.Clear();
      }
      panelListView.View = View.Details;
      panelListView.FullRowSelect = true;
      foreach (Panel panel in panelList)
      {
        ListViewItem item = new ListViewItem(panel.name, 0);
        if (panel.parentDistance == -1)
        {
          item.SubItems.Add("Not Set");
          item.SubItems.Add("Not Set");
        }
        else
        {
          string parent = "";
          foreach (Panel p in panelList)
          {
            if (p.equipId == panel.parentId)
            {
              parent = p.name;
              break;
            }
          }
          item.SubItems.Add(parent);
          item.SubItems.Add(panel.parentDistance.ToString() + "'");
        }
        if (panel.loc.X == 0 && panel.loc.Y == 0)
        {
          item.SubItems.Add("Not Set");
        }
        else
        {
          item.SubItems.Add(
            Math.Round(panel.loc.X / 12, 1).ToString()
              + ", "
              + Math.Round(panel.loc.Y / 12, 1).ToString()
          );
        }
        item.SubItems.Add(panel.equipId);
        item.SubItems.Add(panel.parentId);
        panelListView.Items.Add(item);
        if (!updateOnly)
        {
          filterPanelComboBox.Items.Add(panel.name);
        }
      }
      if (!updateOnly)
      {
        panelListView.Columns.Add("Name", -2, HorizontalAlignment.Left);
        panelListView.Columns.Add("Parent", -2, HorizontalAlignment.Left);
        panelListView.Columns.Add("Parent Distance", -2, HorizontalAlignment.Left);
        panelListView.Columns.Add("Location", -2, HorizontalAlignment.Left);
      }
    }

    private void ResetLocations()
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

      List<string> eqIds = new List<string>();
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
              PooledEquipment eq = new PooledEquipment();
              foreach (DynamicBlockReferenceProperty prop in pc)
              {
                if (prop.PropertyName == "gmep_equip_id" && prop.Value as string != "0")
                {
                  eqIds.Add(prop.Value as string);
                }
              }
            }
          }
          catch { }
        }
      }
      for (int i = 0; i < equipmentList.Count; i++)
      {
        bool found = false;
        foreach (string eqId in eqIds)
        {
          if (eqId == equipmentList[i].equipId)
          {
            found = true;
            break;
          }
        }
        if (!found)
        {
          Equipment eq = equipmentList[i];
          eq.loc = new Point3d(0, 0, 0);
          eq.parentDistance = -1;
          equipmentList[i] = eq;
          gmepDb.UpdateEquipment(eq);
        }
      }
      for (int i = 0; i < panelList.Count; i++)
      {
        bool found = false;
        foreach (string eqId in eqIds)
        {
          if (eqId == panelList[i].equipId)
          {
            found = true;
            break;
          }
        }
        if (!found)
        {
          Panel panel = panelList[i];
          panel.loc = new Point3d(0, 0, 0);
          panel.parentDistance = -1;
          panelList[i] = panel;
          gmepDb.UpdatePanel(panel);
        }
      }
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
                if (prop.PropertyName == "gmep_equip_id" && prop.Value as string == equipId)
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
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        var promptOptions = new PromptPointOptions("\nSelect point for " + equipNo + ": ");
        var promptResult = ed.GetPoint(promptOptions);
        if (promptResult.Status != PromptStatus.OK)
          return new Point3d(0, 0, 0);
        point = promptResult.Value;
        try
        {
          BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
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

            AttributeReference attrRef = new AttributeReference();

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
      List<PooledEquipment> pooledEquipment = new List<PooledEquipment>();
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
              PooledEquipment eq = new PooledEquipment();
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
                pooledEquipment.Add(new PooledEquipment(eq.equipId, eq.parentId, eq.loc));
              }
            }
          }
          catch { }
        }
      }
      for (int i = 0; i < pooledEquipment.Count; i++)
      {
        for (int j = 0; j < pooledEquipment.Count; j++)
        {
          if (pooledEquipment[j].parentId == pooledEquipment[i].equipId)
          {
            PooledEquipment equip = pooledEquipment[j];
            equip.parentDistance = Convert.ToInt32(
              pooledEquipment[j].loc.DistanceTo(pooledEquipment[i].loc) / 12
            );
            pooledEquipment[j] = equip;
          }
        }
      }
      for (int i = 0; i < pooledEquipment.Count; i++)
      {
        bool isMatch = false;
        for (int j = 0; j < equipmentList.Count; j++)
        {
          if (equipmentList[j].equipId == pooledEquipment[i].equipId)
          {
            isMatch = true;
            Equipment equip = equipmentList[j];
            if (
              equip.parentDistance != pooledEquipment[i].parentDistance
              || equip.loc.X != pooledEquipment[i].loc.X
              || equip.loc.Y != pooledEquipment[i].loc.Y
            )
            {
              equip.parentDistance = pooledEquipment[i].parentDistance;
              equip.loc = pooledEquipment[i].loc;
              equipmentList[j] = equip;
              gmepDb.UpdateEquipment(equip);
            }
          }
        }
        if (!isMatch)
        {
          for (int j = 0; j < panelList.Count; j++)
          {
            if (panelList[j].equipId == pooledEquipment[i].equipId)
            {
              Panel panel = panelList[j];
              if (
                panel.parentDistance != pooledEquipment[i].parentDistance
                || panel.loc.X != pooledEquipment[i].loc.X
                || panel.loc.Y != pooledEquipment[i].loc.Y
              )
              {
                panel.parentDistance = pooledEquipment[i].parentDistance;
                panel.loc = pooledEquipment[i].loc;
                panelList[j] = panel;
                gmepDb.UpdatePanel(panel);
              }
            }
          }
        }
      }
      CreatePanelListView(true);
      CreateEquipmentListView(true);
    }

    private void EquipmentListView_MouseDoubleClick(
      object sender,
      System.Windows.Forms.MouseEventArgs e
    )
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
            equipmentList[i] = equipment;
          }
        }
      }
      CalculateDistances();
    }

    private void PanelListView_MouseDoubleClick(
      object sender,
      System.Windows.Forms.MouseEventArgs e
    )
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
        int numSubItems = panelListView.SelectedItems[0].SubItems.Count;
        Point3d p = PlaceEquipment(
          panelListView.SelectedItems[0].SubItems[numSubItems - 2].Text,
          panelListView.SelectedItems[0].SubItems[numSubItems - 1].Text,
          panelListView.SelectedItems[0].Text
        );
      }
      CreatePanelListView(true);
      CalculateDistances();
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
        int numSubItems = 0;
        if (panelListView.SelectedItems.Count > 0)
        {
          numSubItems = panelListView.SelectedItems[0].SubItems.Count;
        }
        Dictionary<string, Point3d> panelLocs = new Dictionary<string, Point3d>();
        bool brk = false;
        foreach (ListViewItem item in panelListView.SelectedItems)
        {
          Point3d p = PlaceEquipment(
            item.SubItems[numSubItems - 2].Text,
            item.SubItems[numSubItems - 1].Text,
            item.Text
          );
          if (p.X == 0 & p.Y == 0 & p.Z == 0)
          {
            brk = true;
            break;
          }
          panelLocs[item.SubItems[numSubItems - 2].Text] = p;
        }
        if (brk || equipmentListView.SelectedItems.Count == 0)
        {
          return;
        }
        numSubItems = equipmentListView.SelectedItems[0].SubItems.Count;
        Dictionary<string, Point3d> equipLocs = new Dictionary<string, Point3d>();
        foreach (ListViewItem item in equipmentListView.SelectedItems)
        {
          Point3d p = PlaceEquipment(
            item.SubItems[numSubItems - 2].Text,
            item.SubItems[numSubItems - 1].Text,
            item.Text
          );
          if (p.X == 0 & p.Y == 0 & p.Z == 0)
          {
            break;
          }
          equipLocs[item.SubItems[numSubItems - 2].Text] = p;
        }
      }
      CalculateDistances();
    }

    private void PlaceAllButton_Click(object sender, EventArgs e)
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
        int numSubItems = 0;
        if (panelListView.Items.Count > 0)
        {
          numSubItems = panelListView.Items[0].SubItems.Count;
        }
        Dictionary<string, Point3d> panelLocs = new Dictionary<string, Point3d>();
        bool brk = false;
        foreach (ListViewItem item in panelListView.Items)
        {
          Point3d p = PlaceEquipment(
            item.SubItems[numSubItems - 2].Text,
            item.SubItems[numSubItems - 1].Text,
            item.Text
          );
          if (p.X == 0 & p.Y == 0 & p.Z == 0)
          {
            brk = true;
            break;
          }
          panelLocs[item.SubItems[numSubItems - 2].Text] = p;
        }
        if (brk || equipmentListView.Items.Count == 0)
        {
          CalculateDistances();
          return;
        }
        numSubItems = equipmentListView.Items[0].SubItems.Count;
        Dictionary<string, Point3d> equipLocs = new Dictionary<string, Point3d>();
        foreach (ListViewItem item in equipmentListView.Items)
        {
          Point3d p = PlaceEquipment(
            item.SubItems[numSubItems - 2].Text,
            item.SubItems[numSubItems - 1].Text,
            item.Text
          );
          if (p.X == 0 & p.Y == 0 & p.Z == 0)
          {
            break;
          }
          equipLocs[item.SubItems[numSubItems - 2].Text] = p;
        }
      }
      CalculateDistances();
    }

    private void RecalculateDistancesButton_Click(object sender, EventArgs e)
    {
      ResetLocations();
      CalculateDistances();
    }

    private void FilterClearButton_Click(object sender, EventArgs e)
    {
      isLoading = true;
      filterPanelComboBox.SelectedIndex = -1;
      filterVoltageComboBox.SelectedIndex = -1;
      filterPhaseComboBox.SelectedIndex = -1;
      filterCategoryComboBox.SelectedIndex = -1;
      filterEquipNoTextBox.Clear();
      isLoading = false;
      CreateEquipmentListView(true);
    }

    private void FilterPanelComboBox_SelectedIndexChanged(object sender, EventArgs e)
    {
      filterPanel = GeneralCommands.GetComboBoxValue(filterPanelComboBox);
      if (!isLoading)
      {
        CreateEquipmentListView(true);
      }
    }

    private void FilterVoltageComboBox_SelectedIndexChanged(object sender, EventArgs e)
    {
      filterVoltage = GeneralCommands.GetComboBoxValue(filterVoltageComboBox);
      if (!isLoading)
      {
        CreateEquipmentListView(true);
      }
    }

    private void FilterPhaseComboBox_SelectedIndexChanged(object sender, EventArgs e)
    {
      filterPhase = GeneralCommands.GetComboBoxValue(filterPhaseComboBox);
      if (!isLoading)
      {
        CreateEquipmentListView(true);
      }
    }

    private void FilterCategoryComboBox_SelectedIndexChanged(object sender, EventArgs e)
    {
      filterCategory = GeneralCommands.GetComboBoxValue(filterCategoryComboBox);
      if (!isLoading)
      {
        CreateEquipmentListView(true);
      }
    }

    private void FilterEquipNoTextBox_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
    {
      if (e.KeyCode == Keys.Enter)
      {
        filterEquipNo = filterEquipNoTextBox.Text.ToUpper();
        CreateEquipmentListView(true);
      }
    }
  }
}
