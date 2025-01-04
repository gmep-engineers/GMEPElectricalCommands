using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using GMEPElectricalCommands.GmepDatabase;

namespace ElectricalCommands.Equipment
{
  public struct Service
  {
    public string id;
    public string name;
    public bool isMultiMeter;
    public int amp;
    public string voltage;

    public Service(string id, string name, string meterConfig, int amp, string voltage)
    {
      this.id = id;
      this.name = name;
      isMultiMeter = meterConfig == "MULTIMETER";
      this.amp = amp;
      this.voltage = voltage;
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
    private List<ListViewItem> transformerListViewList;
    private List<Transformer> transformerList;
    private string projectId;
    private bool isLoading;
    private List<Service> services;
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
      transformerList = gmepDb.GetTransformers(projectId);
      services = gmepDb.GetServices(projectId);
      for (int i = 0; i < panelList.Count; i++)
      {
        foreach (Service service in services)
        {
          if (service.id == panelList[i].parentId)
          {
            Panel panel = panelList[i];
            panel.parentName = service.name;
            panelList[i] = panel;
          }
          else
          {
            bool found = false;
            for (int j = 0; j < panelList.Count; j++)
            {
              if (panelList[i].parentId == panelList[j].id)
              {
                Panel panel = panelList[i];
                panel.parentName = panelList[j].name;
                panelList[i] = panel;
                found = true;
              }
            }
            if (!found)
            {
              for (int j = 0; j < transformerList.Count; j++)
              {
                if (panelList[i].parentId == transformerList[j].id)
                {
                  Panel panel = panelList[i];
                  panel.parentName = transformerList[j].name;
                  panelList[i] = panel;
                }
              }
            }
          }
        }
      }
      for (int i = 0; i < transformerList.Count; i++)
      {
        foreach (Service service in services)
        {
          if (service.id == transformerList[i].parentId)
          {
            Transformer xfmr = transformerList[i];
            xfmr.parentName = service.name;
            transformerList[i] = xfmr;
          }
          else
          {
            for (int j = 0; j < panelList.Count; j++)
            {
              if (transformerList[i].parentId == panelList[j].id)
              {
                Transformer xfmr = transformerList[i];
                xfmr.parentName = panelList[j].name;
                transformerList[i] = xfmr;
              }
            }
          }
        }
      }
      CreateEquipmentListView();
      CreatePanelListView();
      CreateTransformerListView();
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
        if (!String.IsNullOrEmpty(filterEquipNo) && equipment.name != filterEquipNo)
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
        ListViewItem item = new ListViewItem(equipment.name, 0);
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
        item.SubItems.Add(equipment.id);
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
        ContextMenu contextMenu = new ContextMenu();
        contextMenu.MenuItems.Add(
          new MenuItem("Show on plan", new EventHandler(ShowEquipOnPlan_Click))
        );
        equipmentListView.ContextMenu = contextMenu;
      }
    }

    private void ShowEquipOnPlan_Click(object sender, EventArgs e)
    {
      if (equipmentListView.SelectedItems.Count > 0)
      {
        int numSubitems = equipmentListView.SelectedItems[0].SubItems.Count;
        string equipId = equipmentListView.SelectedItems[0].SubItems[numSubitems - 2].Text;
        foreach (Equipment eq in equipmentList)
        {
          if (eq.id == equipId && eq.loc.X != 0 && eq.loc.Y != 0)
          {
            Document doc = Autodesk
              .AutoCAD
              .ApplicationServices
              .Application
              .DocumentManager
              .MdiActiveDocument;

            Editor ed = doc.Editor;
            using (var view = ed.GetCurrentView())
            {
              var UCS2DCS =
                (
                  Matrix3d.Rotation(-view.ViewTwist, view.ViewDirection, view.Target)
                  * Matrix3d.Displacement(view.Target - Point3d.Origin)
                  * Matrix3d.PlaneToWorld(view.ViewDirection)
                ).Inverse() * ed.CurrentUserCoordinateSystem;
              var center = eq.loc.TransformBy(UCS2DCS);
              view.CenterPoint = new Point2d(center.X, center.Y);
              ed.SetCurrentView(view);
            }
            break;
          }
        }
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
        item.SubItems.Add(panel.parentName);
        if (panel.parentDistance == -1)
        {
          item.SubItems.Add("Not Set");
        }
        else
        {
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
        item.SubItems.Add(panel.id);
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
        ContextMenu contextMenu = new ContextMenu();
        contextMenu.MenuItems.Add(
          new MenuItem("Show on plan", new EventHandler(ShowPanelOnPlan_Click))
        );
        panelListView.ContextMenu = contextMenu;
      }
    }

    private void ShowPanelOnPlan_Click(object sender, EventArgs e)
    {
      if (panelListView.SelectedItems.Count > 0)
      {
        int numSubitems = panelListView.SelectedItems[0].SubItems.Count;
        string equipId = panelListView.SelectedItems[0].SubItems[numSubitems - 2].Text;
        foreach (Panel p in panelList)
        {
          if (p.id == equipId && p.loc.X != 0 && p.loc.Y != 0)
          {
            Document doc = Autodesk
              .AutoCAD
              .ApplicationServices
              .Application
              .DocumentManager
              .MdiActiveDocument;

            Editor ed = doc.Editor;
            using (var view = ed.GetCurrentView())
            {
              var UCS2DCS =
                (
                  Matrix3d.Rotation(-view.ViewTwist, view.ViewDirection, view.Target)
                  * Matrix3d.Displacement(view.Target - Point3d.Origin)
                  * Matrix3d.PlaneToWorld(view.ViewDirection)
                ).Inverse() * ed.CurrentUserCoordinateSystem;
              var center = p.loc.TransformBy(UCS2DCS);
              view.CenterPoint = new Point2d(center.X, center.Y);
              ed.SetCurrentView(view);
            }
            break;
          }
        }
      }
    }

    private void CreateTransformerListView(bool updateOnly = false)
    {
      if (updateOnly)
      {
        transformerListView.Items.Clear();
      }
      transformerListView.View = View.Details;
      transformerListView.FullRowSelect = true;
      foreach (Transformer xfmr in transformerList)
      {
        ListViewItem item = new ListViewItem(xfmr.name, 0);
        item.SubItems.Add(xfmr.parentName);
        if (xfmr.parentDistance == -1)
        {
          item.SubItems.Add("Not Set");
        }
        else
        {
          item.SubItems.Add(xfmr.parentDistance.ToString() + "'");
        }
        if (xfmr.loc.X == 0 && xfmr.loc.Y == 0)
        {
          item.SubItems.Add("Not Set");
        }
        else
        {
          item.SubItems.Add(
            Math.Round(xfmr.loc.X / 12, 1).ToString()
              + ", "
              + Math.Round(xfmr.loc.Y / 12, 1).ToString()
          );
        }
        item.SubItems.Add(xfmr.id);
        item.SubItems.Add(xfmr.parentId);
        transformerListView.Items.Add(item);
        if (!updateOnly)
        {
          filterPanelComboBox.Items.Add(xfmr.name);
        }
      }
      if (!updateOnly)
      {
        transformerListView.Columns.Add("Name", -2, HorizontalAlignment.Left);
        transformerListView.Columns.Add("Parent", -2, HorizontalAlignment.Left);
        transformerListView.Columns.Add("Parent Distance", -2, HorizontalAlignment.Left);
        transformerListView.Columns.Add("Location", -2, HorizontalAlignment.Left);
        ContextMenu contextMenu = new ContextMenu();
        contextMenu.MenuItems.Add(
          new MenuItem("Show on plan", new EventHandler(ShowTransformerOnPlan_Click))
        );
        transformerListView.ContextMenu = contextMenu;
      }
    }

    private void ShowTransformerOnPlan_Click(object sender, EventArgs e)
    {
      if (transformerListView.SelectedItems.Count > 0)
      {
        int numSubitems = transformerListView.SelectedItems[0].SubItems.Count;
        string equipId = transformerListView.SelectedItems[0].SubItems[numSubitems - 2].Text;
        foreach (Transformer t in transformerList)
        {
          if (t.id == equipId && t.loc.X != 0 && t.loc.Y != 0)
          {
            Document doc = Autodesk
              .AutoCAD
              .ApplicationServices
              .Application
              .DocumentManager
              .MdiActiveDocument;

            Editor ed = doc.Editor;
            using (var view = ed.GetCurrentView())
            {
              var UCS2DCS =
                (
                  Matrix3d.Rotation(-view.ViewTwist, view.ViewDirection, view.Target)
                  * Matrix3d.Displacement(view.Target - Point3d.Origin)
                  * Matrix3d.PlaneToWorld(view.ViewDirection)
                ).Inverse() * ed.CurrentUserCoordinateSystem;
              var center = t.loc.TransformBy(UCS2DCS);
              view.CenterPoint = new Point2d(center.X, center.Y);
              ed.SetCurrentView(view);
            }
            break;
          }
        }
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
          if (eqId == equipmentList[i].id)
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
          if (eqId == panelList[i].id)
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
      for (int i = 0; i < transformerList.Count; i++)
      {
        bool found = false;
        foreach (string eqId in eqIds)
        {
          if (eqId == transformerList[i].id)
          {
            found = true;
            break;
          }
        }
        if (!found)
        {
          Transformer xfmr = transformerList[i];
          xfmr.loc = new Point3d(0, 0, 0);
          xfmr.parentDistance = -1;
          transformerList[i] = xfmr;
          gmepDb.UpdateTransformer(xfmr);
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
        if (point.X == 0 && point.Y == 0 && point.Z == 0)
        {
          point = new Point3d(0.01, 0, 0);
        }
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
      List<Placeable> pooledEquipment = new List<Placeable>();
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
              Placeable eq = new Placeable();
              foreach (DynamicBlockReferenceProperty prop in pc)
              {
                if (prop.PropertyName == "gmep_equip_id" && prop.Value as string != "0")
                {
                  addEquip = true;
                  eq.id = prop.Value as string;
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
                Placeable p = new Placeable();
                p.id = eq.id;
                p.parentId = eq.parentId;
                p.loc = eq.loc;
                pooledEquipment.Add(p);
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
          if (pooledEquipment[j].parentId == pooledEquipment[i].id)
          {
            Placeable equip = pooledEquipment[j];
            equip.parentDistance =
              Convert.ToInt32(
                Math.Abs(pooledEquipment[j].loc.X - pooledEquipment[i].loc.X)
                  + Math.Abs(pooledEquipment[j].loc.Y - pooledEquipment[i].loc.Y)
              ) / 12;
            pooledEquipment[j] = equip;
          }
        }
      }
      for (int i = 0; i < pooledEquipment.Count; i++)
      {
        bool isMatch = false;
        for (int j = 0; j < equipmentList.Count; j++)
        {
          if (equipmentList[j].id == pooledEquipment[i].id)
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
            if (panelList[j].id == pooledEquipment[i].id)
            {
              isMatch = true;
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
        if (!isMatch)
        {
          for (int j = 0; j < transformerList.Count; j++)
          {
            if (transformerList[j].id == pooledEquipment[i].id)
            {
              isMatch = true;
              Transformer xfmr = transformerList[j];
              if (
                xfmr.parentDistance != pooledEquipment[i].parentDistance
                || xfmr.loc.X != pooledEquipment[i].loc.X
                || xfmr.loc.Y != pooledEquipment[i].loc.Y
              )
              {
                xfmr.parentDistance = pooledEquipment[i].parentDistance;
                xfmr.loc = pooledEquipment[i].loc;
                transformerList[j] = xfmr;
                gmepDb.UpdateTransformer(xfmr);
              }
            }
          }
        }
      }
      CreatePanelListView(true);
      CreateTransformerListView(true);
      CreateEquipmentListView(true);
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
            equipmentList[i].id == equipmentListView.SelectedItems[0].SubItems[numSubItems - 2].Text
          )
          {
            equipment.loc = p;
            equipmentList[i] = equipment;
          }
        }
      }
      CalculateDistances();
    }

    private void PanelListView_MouseDoubleClick(object sender, MouseEventArgs e)
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

    private void TransformerListView_MouseDoubleClick(object sender, MouseEventArgs e)
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
        int numSubItems = transformerListView.SelectedItems[0].SubItems.Count;
        Point3d p = PlaceEquipment(
          transformerListView.SelectedItems[0].SubItems[numSubItems - 2].Text,
          transformerListView.SelectedItems[0].SubItems[numSubItems - 1].Text,
          transformerListView.SelectedItems[0].Text
        );
      }
      CreateTransformerListView(true);
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
        if (
          !brk
          && (
            transformerListView.SelectedItems.Count > 0 || equipmentListView.SelectedItems.Count > 0
          )
        )
        {
          if (transformerListView.SelectedItems.Count > 0)
          {
            numSubItems = transformerListView.SelectedItems[0].SubItems.Count;
            foreach (ListViewItem item in transformerListView.SelectedItems)
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
            }
          }
          if (!brk && equipmentListView.SelectedItems.Count > 0)
          {
            numSubItems = equipmentListView.SelectedItems[0].SubItems.Count;
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
            }
          }
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
        }
        if (!brk && (transformerListView.Items.Count > 0 || equipmentListView.Items.Count > 0))
        {
          if (transformerListView.Items.Count > 0)
          {
            numSubItems = transformerListView.Items[0].SubItems.Count;
            foreach (ListViewItem item in transformerListView.Items)
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
            }
          }
          if (!brk && equipmentListView.Items.Count > 0)
          {
            numSubItems = equipmentListView.Items[0].SubItems.Count;
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
            }
          }
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
      filterEquipNo = "";
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

    private void FilterEquipNoTextBox_KeyUp(object sender, KeyEventArgs e)
    {
      if (e.KeyCode == Keys.Enter)
      {
        filterEquipNo = filterEquipNoTextBox.Text.ToUpper();
        CreateEquipmentListView(true);
      }
    }

    private void MakeSingleLineNodeTreeFromPanel(SLPanel panel)
    {
      foreach (Panel p in panelList)
      {
        if (p.parentId == panel.id)
        {
          SLPanel childPanel = new SLPanel(p.id, p.name, false, false, p.parentDistance);
          childPanel.mainBreakerSize = p.busSize;
          childPanel.voltageSpec = p.voltage;
          if (panel.isDistribution)
          {
            if (!panel.hasMeter)
            {
              childPanel.hasMeter = true;
            }
            childPanel.distributionBreakerSize = p.busSize;
            if (p.busSize >= 400)
            {
              childPanel.hasCts = true;
            }
          }
          if (p.voltage.Contains("3"))
          {
            childPanel.is3Phase = true;
          }
          MakeSingleLineNodeTreeFromPanel(childPanel);
          panel.children.Add(childPanel);
        }
      }
      foreach (Transformer t in transformerList)
      {
        if (t.parentId == panel.id)
        {
          SLDisconnect disc = new SLDisconnect(t.name);
          bool is3Phase = false;
          if (t.voltageSpec.Contains("3"))
          {
            is3Phase = true;
          }
          double voltage = 208;
          if (t.voltageSpec.StartsWith("480"))
          {
            voltage = 480;
          }
          double amperage = t.kva * 1000 / voltage;
          int mainBreakerSize = 0;
          string grounding = "(N)";
          switch (amperage)
          {
            case var _ when amperage <= 100:
              mainBreakerSize = 100;
              grounding += "3/4\"C. (1#8CU.)";
              break;
            case var _ when amperage <= 125:
              mainBreakerSize = 125;
              grounding += "3/4\"C. (1#8CU.)";
              break;
            case var _ when amperage <= 150:
              mainBreakerSize = 150;
              grounding += "3/4\"C. (1#8CU.)";
              break;
            case var _ when amperage <= 175:
              mainBreakerSize = 175;
              grounding += "3/4\"C. (1#8CU.)";
              break;
            case var _ when amperage <= 200:
              mainBreakerSize = 200;
              grounding += "3/4\"C. (1#6CU.)";
              break;
            case var _ when amperage <= 225:
              mainBreakerSize = 225;
              grounding += "3/4\"C. (1#6CU.)";
              break;
            case var _ when amperage <= 250:
              mainBreakerSize = 250;
              grounding += "3/4\"C. (1#6CU.)";
              break;
            case var _ when amperage <= 275:
              mainBreakerSize = 275;
              grounding += "3/4\"C. (1#6CU.)";
              break;
            case var _ when amperage <= 400:
              mainBreakerSize = 400;
              grounding += "3/4\"C. (1#3CU.)";
              break;
            case var _ when amperage <= 500:
              mainBreakerSize = 500;
              grounding += "3/4\"C. (1#2CU.)";
              break;
            case var _ when amperage <= 600:
              mainBreakerSize = 600;
              grounding += "3/4\"C. (1#1CU.)";
              break;
            case var _ when amperage <= 800:
              mainBreakerSize = 800;
              grounding += "3/4\"C. (1#1/0CU.)";
              break;
          }
          disc.is3Phase = is3Phase;
          disc.voltage = voltage;
          disc.mainBreakerSize = mainBreakerSize;
          disc.fromDistribution = panel.isDistribution;
          disc.parentDistance = t.parentDistance;
          if (panel.isDistribution && !panel.hasMeter)
          {
            disc.hasMeter = true;
            if (mainBreakerSize > 200)
            {
              disc.hasCts = true;
            }
          }
          SLTransformer childXfmr = new SLTransformer(t.id, t.name);
          childXfmr.voltage = voltage;
          childXfmr.parentDistance = 0;
          childXfmr.is3Phase = panel.is3Phase;
          childXfmr.mainBreakerSize = mainBreakerSize;
          childXfmr.grounding = grounding;
          disc.children.Add(childXfmr);
          MakeSingleLineNodeTreeFromTransformer(childXfmr);
          panel.children.Add(disc);
        }
      }
    }

    private void MakeSingleLineNodeTreeFromTransformer(SLTransformer transformer)
    {
      foreach (Panel p in panelList)
      {
        if (p.parentId == transformer.id)
        {
          SLPanel childPanel = new SLPanel(p.id, p.name, false, false, p.parentDistance);
          MakeSingleLineNodeTreeFromPanel(childPanel);
          transformer.children.Add(childPanel);
        }
      }
    }

    private void MakeSingleLineNodeTreeFromService(SLServiceFeeder sf)
    {
      foreach (Panel panel in panelList)
      {
        if (panel.parentId == sf.id)
        {
          bool hasMeter = false;
          if (!sf.isMultiMeter)
          {
            hasMeter = true;
          }
          SLPanel p = new SLPanel(panel.id, panel.name, true, hasMeter, panel.parentDistance);
          if (sf.amp >= 400)
          {
            p.hasCts = true;
          }
          if (sf.amp >= 1200)
          {
            p.hasGfp = true;
          }
          p.distributionBreakerSize = sf.amp;
          p.voltageSpec = sf.voltage;
          p.is3Phase = false;
          if (sf.voltage.Contains("3"))
          {
            p.is3Phase = true;
          }
          MakeSingleLineNodeTreeFromPanel(p);
          sf.children.Add(p);
        }
      }
    }

    private SingleLine MakeSingleLineNodeTree()
    {
      SingleLine singleLine = new SingleLine();
      foreach (Service service in services)
      {
        SLServiceFeeder sf = new SLServiceFeeder(
          service.id,
          service.name,
          service.isMultiMeter,
          service.amp,
          service.voltage
        );
        MakeSingleLineNodeTreeFromService(sf);
        singleLine.children.Add(sf);
      }
      return singleLine;
    }

    private void MakeSingleLineButton_Click(object sender, EventArgs e)
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
        Point3d startingPoint;
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
          var promptOptions = new PromptPointOptions("\nSelect upper left point:");
          var promptResult = ed.GetPoint(promptOptions);
          if (promptResult.Status == PromptStatus.OK)
            startingPoint = promptResult.Value;
          else
          {
            return;
          }
        }
        SingleLine singleLineNodeTree = MakeSingleLineNodeTree();
        singleLineNodeTree.AggregateWidths();
        singleLineNodeTree.SetChildStartingPoints(startingPoint);
        singleLineNodeTree.Make();
      }
    }

    private void CreateEquipmentSchedule(Document doc, Database db, Editor ed, Point3d startPoint)
    {
      var spaceId =
        (db.TileMode == true)
          ? SymbolUtilityServices.GetBlockModelSpaceId(db)
          : SymbolUtilityServices.GetBlockPaperSpaceId(db);
      using (var tr = db.TransactionManager.StartTransaction())
      {
        var btr = (BlockTableRecord)tr.GetObject(spaceId, OpenMode.ForWrite);
        Table tb = new Table();
        tb.TableStyle = db.Tablestyle;
        tb.Position = startPoint;
        int tableRows = equipmentList.Count + 3;
        int tableCols = 10;
        tb.SetSize(tableRows, tableCols);
        tb.SetRowHeight(0.25);
        tb.Cells[0, 0].TextString = "ELECTRICAL EQUIPMENT SCHEDULE";
        CellRange range = CellRange.Create(tb, 1, 0, 2, 0);
        tb.MergeCells(range);
        range = CellRange.Create(tb, 1, 1, 2, 1);
        tb.MergeCells(range);
        range = CellRange.Create(tb, 1, 2, 1, 6);
        tb.MergeCells(range);
        range = CellRange.Create(tb, 1, 7, 1, 9);
        tb.MergeCells(range);
        var textStyleId = PanelCommands.GetTextStyleId("gmep");
        tb.Layer = "E-TXT1";
        tb.Cells[1, 0].TextString = "TAG";
        tb.Cells[1, 1].TextString = "DESCRIPTION";
        tb.Cells[1, 2].TextString = "ELECTRICAL";
        tb.Cells[2, 2].TextString = "VOLT.";
        tb.Cells[2, 3].TextString = "FLA";
        tb.Cells[2, 4].TextString = "HP";
        tb.Cells[2, 5].TextString = "MCA";
        tb.Cells[2, 6].TextString = "PH.";
        tb.Columns[0].Width = 0.75;
        tb.Columns[1].Width = 2.25;
        tb.Columns[2].Width = 0.5;
        tb.Columns[3].Width = 0.5;
        tb.Columns[4].Width = 0.5;
        tb.Columns[5].Width = 0.5;
        tb.Columns[6].Width = 0.5;
        tb.Columns[7].Width = 1.3;
        tb.Columns[8].Width = 0.7;
        tb.Columns[8].Width = 1.3;
        tb.Cells[1, 7].TextString = "ROUGH-IN";
        tb.Cells[2, 7].TextString = "CONNECTION";
        tb.Cells[2, 8].TextString = "HEIGHT";
        tb.Cells[2, 9].TextString = "WIRE SIZE";
        for (int i = 0; i < tableRows; i++)
        {
          for (int j = 0; j < tableCols; j++)
          {
            tb.Cells[i, j].TextStyleId = textStyleId;
            tb.Cells[i, j].TextHeight = (0.0832);
            tb.Cells[i, j].Alignment = CellAlignment.MiddleCenter;
          }
        }
        for (int i = 0; i < equipmentList.Count; i++)
        {
          int row = i + 3;
          tb.Cells[row, 0].TextString = equipmentList[i].name;
          tb.Cells[row, 1].TextString = equipmentList[i].description.ToUpper();
          tb.Cells[row, 2].TextString = equipmentList[i].voltage.ToString();
          tb.Cells[row, 3].TextString =
            equipmentList[i].fla > 0 ? Math.Round(equipmentList[i].fla, 1).ToString() : "-";

          tb.Cells[row, 6].TextString = equipmentList[i].is3Phase ? "3" : "1";
        }
        BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
        btr.AppendEntity(tb);
        tr.AddNewlyCreatedDBObject(tb, true);
        tr.Commit();
      }
    }

    private void CreateEquipmentSchedule_Click(object sender, EventArgs e)
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
        Point3d startingPoint;
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
          var promptOptions = new PromptPointOptions("\nSelect upper left point:");
          var promptResult = ed.GetPoint(promptOptions);
          if (promptResult.Status == PromptStatus.OK)
            startingPoint = promptResult.Value;
          else
          {
            return;
          }
        }
        CreateEquipmentSchedule(doc, db, ed, startingPoint);
      }
    }
  }
}
