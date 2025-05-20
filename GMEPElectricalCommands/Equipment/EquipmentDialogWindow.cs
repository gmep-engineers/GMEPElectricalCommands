using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2010.Excel;
using ElectricalCommands.ElectricalEntity;
using ElectricalCommands.Lighting;
using Emgu.CV.ML;
using GMEPElectricalCommands.GmepDatabase;
using Table = Autodesk.AutoCAD.DatabaseServices.Table;

namespace ElectricalCommands.Equipment
{
  public enum EquipmentType
  {
    Duplex,
    JBox,
    Disconnect,
    Panel,
    Transformer,
  }

  public partial class EquipmentDialogWindow : Form
  {
    private string filterPanel;
    private string filterVoltage;
    private string filterPhase;
    private string filterEquipNo;
    private string filterCategory;
    private List<ListViewItem> equipmentListViewList;
    private List<ElectricalEntity.Equipment> equipmentList;
    private List<ListViewItem> panelListViewList;
    private List<ElectricalEntity.Panel> panelList;
    private List<ListViewItem> transformerListViewList;
    private List<ElectricalEntity.Transformer> transformerList;
    private string projectId;
    private bool isLoading;
    private List<ElectricalEntity.Service> services;
    public GmepDatabase gmepDb = new GmepDatabase();

    public EquipmentDialogWindow(EquipmentCommands EquipCommands)
    {
      InitializeComponent();
    }

    public void InitializeModal()
    {
      projectId = gmepDb.GetProjectId(CADObjectCommands.GetProjectNoFromFileName());
      panelList = gmepDb.GetPanels(projectId);
      equipmentList = gmepDb.GetEquipment(projectId);
      transformerList = gmepDb.GetTransformers(projectId);
      services = gmepDb.GetServices(projectId);
      for (int i = 0; i < panelList.Count; i++)
      {
        foreach (ElectricalEntity.Service service in services)
        {
          if (service.Id == panelList[i].ParentId)
          {
            ElectricalEntity.Panel panel = panelList[i];
            panel.ParentName = service.Name;
            panelList[i] = panel;
          }
          else
          {
            bool found = false;
            for (int j = 0; j < panelList.Count; j++)
            {
              if (panelList[i].ParentId == panelList[j].Id)
              {
                ElectricalEntity.Panel panel = panelList[i];
                panel.ParentName = panelList[j].Name;
                panelList[i] = panel;
                found = true;
              }
            }
            if (!found)
            {
              for (int j = 0; j < transformerList.Count; j++)
              {
                if (panelList[i].ParentId == transformerList[j].Id)
                {
                  ElectricalEntity.Panel panel = panelList[i];
                  panel.ParentName = transformerList[j].Name;
                  panelList[i] = panel;
                }
              }
            }
          }
        }
      }
      for (int i = 0; i < transformerList.Count; i++)
      {
        foreach (ElectricalEntity.Service service in services)
        {
          if (service.Id == transformerList[i].ParentId)
          {
            ElectricalEntity.Transformer xfmr = transformerList[i];
            xfmr.ParentName = service.Name;
            transformerList[i] = xfmr;
          }
          else
          {
            for (int j = 0; j < panelList.Count; j++)
            {
              if (transformerList[i].ParentId == panelList[j].Id)
              {
                ElectricalEntity.Transformer xfmr = transformerList[i];
                xfmr.ParentName = panelList[j].Name;
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
      foreach (ElectricalEntity.Equipment equipment in equipmentList)
      {
        if (
          !String.IsNullOrEmpty(filterPanel)
          && equipment.ParentName.ToUpper().Replace("PANEL", "").Trim() != filterPanel
        )
        {
          continue;
        }
        if (!String.IsNullOrEmpty(filterVoltage) && equipment.Voltage.ToString() != filterVoltage)
        {
          continue;
        }
        if (
          !String.IsNullOrEmpty(filterPhase)
          && filterPhase.ToString() == "1"
          && equipment.Is3Phase
        )
        {
          continue;
        }
        if (
          !String.IsNullOrEmpty(filterPhase)
          && filterPhase.ToString() == "3"
          && !equipment.Is3Phase
        )
        {
          continue;
        }
        if (!String.IsNullOrEmpty(filterEquipNo) && equipment.Name != filterEquipNo)
        {
          continue;
        }
        if (
          !String.IsNullOrEmpty(filterCategory)
          && equipment.Category.ToUpper() != filterCategory.ToUpper()
        )
        {
          continue;
        }
        ListViewItem item = new ListViewItem(equipment.Name, 0);
        item.SubItems.Add(equipment.Description);
        item.SubItems.Add(equipment.Category);
        item.SubItems.Add(equipment.ParentName.ToUpper().Replace("PANEL", "").Trim());
        item.SubItems.Add(equipment.Circuit.ToString());
        if (equipment.ParentDistance == -1)
        {
          item.SubItems.Add("Not Set");
        }
        else
        {
          item.SubItems.Add(equipment.ParentDistance.ToString() + "'");
        }
        item.SubItems.Add(equipment.Voltage.ToString());
        item.SubItems.Add(equipment.Is3Phase ? "3" : "1");
        if (equipment.Location.X == 0 && equipment.Location.Y == 0)
        {
          item.SubItems.Add("Not Set");
        }
        else
        {
          item.SubItems.Add(
            Math.Round(equipment.Location.X / 12, 1).ToString()
              + ", "
              + Math.Round(equipment.Location.Y / 12, 1).ToString()
          );
        }
        item.SubItems.Add(equipment.Category);
        item.SubItems.Add(equipment.ConnectionSymbol);
        item.SubItems.Add(equipment.IsHidden.ToString());
        item.SubItems.Add(equipment.Id);
        item.SubItems.Add(equipment.ParentId);

        equipmentListView.Items.Add(item);
      }
      if (!updateOnly)
      {
        equipmentListView.Columns.Add("Equip #", -2, HorizontalAlignment.Left);
        equipmentListView.Columns.Add("Description", -2, HorizontalAlignment.Left);
        equipmentListView.Columns.Add("Category", -2, HorizontalAlignment.Left);
        equipmentListView.Columns.Add("Panel", -2, HorizontalAlignment.Left);
        equipmentListView.Columns.Add("Circuit", -2, HorizontalAlignment.Left);
        equipmentListView.Columns.Add("Panel Distance", -2, HorizontalAlignment.Left);
        equipmentListView.Columns.Add("Voltage", -2, HorizontalAlignment.Left);
        equipmentListView.Columns.Add("Phase", -2, HorizontalAlignment.Left);
        equipmentListView.Columns.Add("Location", -2, HorizontalAlignment.Left);
        equipmentListView.Columns.Add("Category", -2, HorizontalAlignment.Left);
        equipmentListView.Columns.Add("Connection Symbol", -2, HorizontalAlignment.Left);
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
        foreach (ElectricalEntity.Equipment eq in equipmentList)
        {
          if (eq.Id == equipId && eq.Location.X != 0 && eq.Location.Y != 0)
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
              var center = eq.Location.TransformBy(UCS2DCS);
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
      foreach (ElectricalEntity.Panel panel in panelList)
      {
        ListViewItem item = new ListViewItem(panel.Name, 0);
        item.SubItems.Add(panel.ParentName);
        if (panel.ParentDistance == -1)
        {
          item.SubItems.Add("Not Set");
        }
        else
        {
          item.SubItems.Add(panel.ParentDistance.ToString() + "'");
        }
        if (panel.Location.X == 0 && panel.Location.Y == 0)
        {
          item.SubItems.Add("Not Set");
        }
        else
        {
          item.SubItems.Add(
            Math.Round(panel.Location.X / 12, 1).ToString()
              + ", "
              + Math.Round(panel.Location.Y / 12, 1).ToString()
          );
        }
        item.SubItems.Add(panel.IsHidden.ToString());
        item.SubItems.Add(panel.Id);
        item.SubItems.Add(panel.ParentId);
        panelListView.Items.Add(item);
        if (!updateOnly)
        {
          filterPanelComboBox.Items.Add(panel.Name);
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
        foreach (ElectricalEntity.Panel p in panelList)
        {
          if (p.Id == equipId && p.Location.X != 0 && p.Location.Y != 0)
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
              var center = p.Location.TransformBy(UCS2DCS);
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
        ListViewItem item = new ListViewItem(xfmr.Name, 0);
        item.SubItems.Add(xfmr.ParentName);
        if (xfmr.ParentDistance == -1)
        {
          item.SubItems.Add("Not Set");
        }
        else
        {
          item.SubItems.Add(xfmr.ParentDistance.ToString() + "'");
        }
        if (xfmr.Location.X == 0 && xfmr.Location.Y == 0)
        {
          item.SubItems.Add("Not Set");
        }
        else
        {
          item.SubItems.Add(
            Math.Round(xfmr.Location.X / 12, 1).ToString()
              + ", "
              + Math.Round(xfmr.Location.Y / 12, 1).ToString()
          );
        }
        item.SubItems.Add(xfmr.IsHidden.ToString());
        item.SubItems.Add(xfmr.Id);
        item.SubItems.Add(xfmr.ParentId);
        transformerListView.Items.Add(item);
        if (!updateOnly)
        {
          filterPanelComboBox.Items.Add(xfmr.Name);
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
          if (t.Id == equipId && t.Location.X != 0 && t.Location.Y != 0)
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
              var center = t.Location.TransformBy(UCS2DCS);
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
          if (eqId == equipmentList[i].Id)
          {
            found = true;
            break;
          }
        }
        if (!found)
        {
          ElectricalEntity.Equipment eq = equipmentList[i];
          eq.Location = new Point3d(0, 0, 0);
          if (!equipmentList[i].IsHidden)
          {
            int parentDistance = -1;
            foreach (ElectricalEntity.Panel p in panelList)
            {
              if (p.Id == equipmentList[i].ParentId)
              {
                if (p.IsHidden)
                {
                  parentDistance = equipmentList[i].ParentDistance;
                }
              }
            }
            foreach (Transformer t in transformerList)
            {
              if (t.Id == equipmentList[i].ParentId)
              {
                if (t.IsHidden)
                {
                  parentDistance = equipmentList[i].ParentDistance;
                }
              }
            }
            eq.ParentDistance = parentDistance;
          }
          else
          {
            eq.ParentDistance = equipmentList[i].ParentDistance;
          }
          equipmentList[i] = eq;
          gmepDb.UpdateEquipment(eq);
        }
      }
      for (int i = 0; i < panelList.Count; i++)
      {
        bool found = false;
        foreach (string eqId in eqIds)
        {
          if (eqId == panelList[i].Id)
          {
            found = true;
            break;
          }
        }
        if (!found)
        {
          ElectricalEntity.Panel panel = panelList[i];
          panel.Location = new Point3d(0, 0, 0);
          if (!panelList[i].IsHidden)
          {
            int parentDistance = -1;
            foreach (ElectricalEntity.Panel p in panelList)
            {
              if (p.Id == panelList[i].ParentId)
              {
                if (p.IsHidden)
                {
                  parentDistance = panelList[i].ParentDistance;
                }
              }
            }
            foreach (Transformer t in transformerList)
            {
              if (t.Id == panelList[i].ParentId)
              {
                if (t.IsHidden)
                {
                  parentDistance = panelList[i].ParentDistance;
                }
              }
            }
            panel.ParentDistance = parentDistance;
          }
          else
          {
            panel.ParentDistance = panelList[i].ParentDistance;
          }
          panelList[i] = panel;
          gmepDb.UpdatePanel(panel);
        }
      }
      for (int i = 0; i < transformerList.Count; i++)
      {
        bool found = false;
        foreach (string eqId in eqIds)
        {
          if (eqId == transformerList[i].Id)
          {
            found = true;
            break;
          }
        }
        if (!found)
        {
          Transformer xfmr = transformerList[i];
          xfmr.Location = new Point3d(0, 0, 0);
          if (!transformerList[i].IsHidden)
          {
            int parentDistance = -1;
            foreach (ElectricalEntity.Panel p in panelList)
            {
              if (p.Id == transformerList[i].ParentId)
              {
                if (p.IsHidden)
                {
                  parentDistance = transformerList[i].ParentDistance;
                }
              }
            }

            xfmr.ParentDistance = parentDistance;
          }
          else
          {
            xfmr.ParentDistance = transformerList[i].ParentDistance;
          }
          transformerList[i] = xfmr;
          gmepDb.UpdateTransformer(xfmr);
        }
      }
    }

    private void PlaceConvenienceReceptacles(
      string equipId,
      string parentId,
      string equipNo,
      string circuitNo,
      string connectionSymbol
    )
    {
      int duplexCount = gmepDb.GetNumDuplex(equipId);
      if (duplexCount == 0)
      {
        return;
      }
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Core
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Editor ed = doc.Editor;
      Database db = doc.Database;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        LayerTable acLyrTbl;
        acLyrTbl = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

        string sLayerName = "E-SYM1";

        if (acLyrTbl.Has(sLayerName) == true)
        {
          db.Clayer = acLyrTbl[sLayerName];
          tr.Commit();
        }
      }
      Dictionary<string, int> duplexDict = LightingDialogWindow.GetNumObjectsOnPlan(
        "gmep_equip_id"
      );

      int currentNumDuplexes = 0;

      if (duplexDict.ContainsKey(equipId))
      {
        currentNumDuplexes = duplexDict[equipId];
      }
      for (int i = currentNumDuplexes; i < duplexCount; i++)
      {
        ed.WriteMessage(
          "\nPlace " + (i + 1).ToString() + "/" + duplexCount + " for '" + equipNo + "'"
        );
        ObjectId blockId;
        try
        {
          Point3d point;
          double rotation = 0;
          using (Transaction tr = db.TransactionManager.StartTransaction())
          {
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord block = (BlockTableRecord)
              tr.GetObject(bt[connectionSymbol], OpenMode.ForRead);
            BlockJig blockJig = new BlockJig();

            PromptResult res = blockJig.DragMe(block.ObjectId, out point);

            if (res.Status == PromptStatus.OK)
            {
              BlockTableRecord curSpace = (BlockTableRecord)
                tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

              BlockReference br = new BlockReference(point, block.ObjectId);
              RotateJig rotateJig = new RotateJig(br);
              PromptResult rotatePromptResult = ed.Drag(rotateJig);

              if (rotatePromptResult.Status != PromptStatus.OK)
              {
                return;
              }
              rotation = br.Rotation;

              curSpace.AppendEntity(br);

              tr.AddNewlyCreatedDBObject(br, true);
              blockId = br.Id;
              double circuitOffsetX = 0;
              double circuitOffsetY = 0;
              switch (rotation)
              {
                case var _ when rotation > 5.49:
                  circuitOffsetY = -4.5;
                  circuitOffsetX = 4.5;
                  break;
                case var _ when rotation > 4.71:
                  circuitOffsetY = -4.5;
                  break;
                case var _ when rotation > 2.35:
                  circuitOffsetY = 4.5;
                  circuitOffsetX = -4.5;
                  break;
                case var _ when rotation > 1.57:
                  circuitOffsetX = 4.5;
                  circuitOffsetY = -4.5;
                  break;
                default:
                  circuitOffsetY = -4.5;
                  circuitOffsetX = 4.5;
                  break;
              }
              circuitOffsetY = circuitOffsetY * 0.25 / CADObjectCommands.Scale;
              circuitOffsetX = circuitOffsetX * 0.25 / CADObjectCommands.Scale;
              GeneralCommands.CreateAndPositionText(
                tr,
                circuitNo,
                "gmep",
                0.0938 * 12.0 / CADObjectCommands.Scale,
                0.85,
                2,
                "E-TXT1",
                new Point3d(point.X - (circuitOffsetX), point.Y - (circuitOffsetY), 0)
              );
              DynamicBlockReferencePropertyCollection pc =
                br.DynamicBlockReferencePropertyCollection;

              foreach (DynamicBlockReferenceProperty prop in pc)
              {
                if (prop.PropertyName == "gmep_equip_id" && prop.Value as string == "0")
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
                  prop.Value = "true";
                }
              }
              tr.Commit();
            }
          }
        }
        catch (Exception ex)
        {
          Console.WriteLine(ex.ToString());
        }
      }
    }

    private Point3d? PlaceEquipment(
      string equipId,
      string parentId,
      string equipNo,
      EquipmentType equipType,
      string category,
      string circuitNo,
      string connectionSymbol
    )
    {
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

      //editing the connectionSymbol string
      if (connectionSymbol.Contains(" W/ "))
      {
        string[] parts = connectionSymbol.Split(new string[] { " W/ " }, StringSplitOptions.None);

        if (parts.Length == 2)
        {
          string beforeW = parts[0];
          string afterW = parts[1];
          connectionSymbol = beforeW + afterW;
        }
      }
      connectionSymbol = "GMEP " + connectionSymbol.ToUpper();

      if (category.ToLower().StartsWith("convenience"))
      {
        PlaceConvenienceReceptacles(equipId, parentId, equipNo, circuitNo, connectionSymbol);
        return new Point3d();
      }
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
      Point3d firstClickPoint;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        Point3d point;
        double rotation = 0;

        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord block =
          tr.GetObject(bt[connectionSymbol], OpenMode.ForRead) as BlockTableRecord;

        BlockJig blockJig = new BlockJig();

        PromptResult res = blockJig.DragMe(block.ObjectId, out point);
        firstClickPoint = point;
        if (res.Status == PromptStatus.OK)
        {
          BlockTableRecord curSpace = (BlockTableRecord)
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForRead);
        }

        BlockReference br = new BlockReference(Point3d.Origin, block.ObjectId);
        RotateJig rotateJig = new RotateJig(br);
        PromptResult blockPromptResult = ed.Drag(rotateJig);
        double scaleFactor = 0.25;
        if (equipType != EquipmentType.Panel && equipType != EquipmentType.Transformer)
        {
          scaleFactor = scale;
        }
        if (blockPromptResult.Status == PromptStatus.OK)
        {
          BlockTableRecord currentSpace =
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
          currentSpace.AppendEntity(br);
          tr.AddNewlyCreatedDBObject(br, true);
          br.ScaleFactors = new Scale3d(0.25 / scaleFactor);
          br.Layer = "E-SYM1";
          rotation = br.Rotation;

          if (br.IsDynamicBlock)
          {
            DynamicBlockReferencePropertyCollection pc = br.DynamicBlockReferencePropertyCollection;
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
              if (prop.PropertyName == "gmep_equip_no")
              {
                prop.Value = equipNo;
              }
              if (prop.PropertyName == "gmep_equip_locator")
              {
                prop.Value = "true";
              }
            }
          }

          //Setting attributes if blockname is j-box or j-boxswitch
          if (connectionSymbol == "GMEP J-BOX" || connectionSymbol == "GMEP J-BOXSWITCH")
          {
            foreach (ObjectId objId in block)
            {
              DBObject obj = tr.GetObject(objId, OpenMode.ForRead);
              AttributeDefinition attDef = obj as AttributeDefinition;
              if (attDef != null && !attDef.Constant)
              {
                using (AttributeReference attRef = new AttributeReference())
                {
                  attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                  if (attRef.Tag == "J")
                  {
                    attRef.TextString = "J";
                    attRef.Rotation = attRef.Rotation - rotation;
                  }
                  br.AttributeCollection.AppendAttribute(attRef);
                  tr.AddNewlyCreatedDBObject(attRef, true);
                }
              }
            }
          }
        }
        double labelOffsetX = 0;
        double labelOffsetY = 0;
        switch (equipType)
        {
          case EquipmentType.Duplex:
            switch (rotation)
            {
              case var _ when rotation > 5.49:
                labelOffsetY = -4.5;
                break;
              case var _ when rotation > 4.71:
                labelOffsetX = -4.5;
                break;
              case var _ when rotation > 2.35:
                labelOffsetY = 4.5;
                break;
              case var _ when rotation > 1.57:
                labelOffsetX = 4.5;
                break;
              default:
                labelOffsetY = -4.5;
                break;
            }
            break;
          case EquipmentType.Panel:
            break;
          // TODO check for remaining connection types
        }
        labelOffsetY = labelOffsetY * 0.25 / scale;
        labelOffsetX = labelOffsetX * 0.25 / scale;
        firstClickPoint = new Point3d(
          firstClickPoint.X + labelOffsetX,
          firstClickPoint.Y + labelOffsetY,
          0
        );

        if (equipType != EquipmentType.Panel && equipType != EquipmentType.Transformer)
        {
          double circuitOffsetX = 0;
          double circuitOffsetY = 0;
          switch (rotation)
          {
            case var _ when rotation > 5.49:
              circuitOffsetY = 4.5;
              circuitOffsetX = 4.5;
              break;
            case var _ when rotation > 4.71:
              circuitOffsetY = -4.5;
              break;
            case var _ when rotation > 2.35:
              circuitOffsetY = 4.5;
              circuitOffsetX = -4.5;
              break;
            case var _ when rotation > 1.57:
              circuitOffsetX = 4.5;
              circuitOffsetY = -4.5;
              break;
            default:
              circuitOffsetY = -4.5;
              circuitOffsetX = 4.5;
              break;
          }
          circuitOffsetY = circuitOffsetY * 0.25 / scale;
          circuitOffsetX = circuitOffsetX * 0.25 / scale;
          GeneralCommands.CreateAndPositionText(
            tr,
            circuitNo,
            "gmep",
            0.0938 * 12.0 / CADObjectCommands.Scale,
            0.85,
            2,
            "E-TXT1",
            new Point3d(
              firstClickPoint.X - (circuitOffsetX),
              firstClickPoint.Y - (circuitOffsetY),
              0
            )
          );
        }
        tr.Commit();
      }
      // Placing down a switch for a j-box switch
      if (connectionSymbol.ToLower().Contains("switch"))
      {
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          Point3d point;
          BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
          BlockTableRecord block =
            tr.GetObject(bt["GMEP J-BOXSWITCHOBJ"], OpenMode.ForRead) as BlockTableRecord;

          BlockJig blockJig = new BlockJig();

          PromptResult res = blockJig.DragMe(block.ObjectId, out point);

          if (res.Status == PromptStatus.OK)
          {
            BlockTableRecord curSpace = (BlockTableRecord)
              tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            BlockReference br = new BlockReference(point, block.ObjectId);

            RotateJig rotateJig = new RotateJig(br);
            PromptResult rotatePromptResult = ed.Drag(rotateJig);

            if (rotatePromptResult.Status != PromptStatus.OK)
            {
              return null;
            }

            curSpace.AppendEntity(br);

            tr.AddNewlyCreatedDBObject(br, true);
          }

          tr.Commit();
        }
      }

      {
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
        Point3d textAlignmentReferencePoint = thirdClickOccurred
          ? thirdClickPoint
          : secondClickPoint;
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
            ObjectId blockId = bt["EQUIP_MARKER"];
            using (BlockReference acBlkRef = new BlockReference(labelInsertionPoint, blockId))
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
              AttributeDefinition attrDef = new AttributeDefinition();
              attrDef.Position = labelInsertionPoint;
              attrDef.LockPositionInBlock = true;
              attrDef.Tag = equipNo;
              attrDef.IsMTextAttributeDefinition = false;
              attrDef.TextString = equipNo;
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
      }
      if (connectionSymbol.Contains("DISCONNECT"))
      {
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
        Point3d textAlignmentReferencePoint = thirdClickOccurred
          ? thirdClickPoint
          : secondClickPoint;
        Point3d comparisonPoint = thirdClickOccurred ? secondClickPoint : firstClickPoint;

        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
        {
          BlockTableRecord btr = (BlockTableRecord)
            tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite);

          // Create MText object with dynamic text height based on scale
          double textHeight = 1.125 / scale;
          MText mText = new MText();
          mText.SetDatabaseDefaults();
          mText.TextHeight = textHeight;

          // Retrieve the "rpm" text style
          using (Transaction tr2 = doc.Database.TransactionManager.StartTransaction())
          {
            TextStyleTable textStyleTable = (TextStyleTable)
              tr2.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead);
            if (textStyleTable.Has("rpm"))
            {
              mText.TextStyleId = textStyleTable["rpm"];
            }
            else
            {
              ed.WriteMessage("\nText style 'rpm' not found. Using default text style.");
              mText.TextStyleId = doc.Database.Textstyle;
            }
            tr2.Commit();
          }

          mText.Width = 0;
          mText.Layer = "E-TXT1";

          // Determine justification and position
          if (textAlignmentReferencePoint.X > comparisonPoint.X)
          {
            mText.Attachment = AttachmentPoint.TopLeft;
            mText.Location = new Point3d(
              textAlignmentReferencePoint.X + 0.25 / scale,
              textAlignmentReferencePoint.Y + textHeight / 2,
              textAlignmentReferencePoint.Z
            );
          }
          else
          {
            mText.Attachment = AttachmentPoint.TopRight;
            mText.Location = new Point3d(
              textAlignmentReferencePoint.X - 0.25 / scale,
              textAlignmentReferencePoint.Y + textHeight / 2,
              textAlignmentReferencePoint.Z
            );
          }
          (double fla, int voltage, bool is3Phase) = gmepDb.GetEquipmentPowerSpecs(equipId);
          string disconnectSize = CADObjectCommands.GetConnectionTypeFromFlaVoltage(
            fla,
            voltage,
            false,
            is3Phase
          );
          mText.Contents = disconnectSize.Contains("AS")
            ? disconnectSize + "\nDISCONNECT"
            : disconnectSize;
          btr.AppendEntity(mText);
          tr.AddNewlyCreatedDBObject(mText, true);

          tr.Commit();
        }
      }
      return firstClickPoint;
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
      List<PlaceableElectricalEntity> pooledEquipment = new List<PlaceableElectricalEntity>();
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
              PlaceableElectricalEntity eq = new PlaceableElectricalEntity();
              foreach (DynamicBlockReferenceProperty prop in pc)
              {
                if (prop.PropertyName == "gmep_equip_locator" && prop.Value as string == "true")
                {
                  addEquip = true;
                }
                if (prop.PropertyName == "gmep_equip_id" && prop.Value as string != "0")
                {
                  eq.Id = prop.Value as string;
                }
                if (prop.PropertyName == "gmep_equip_parent_id" && prop.Value as string != "0")
                {
                  eq.ParentId = prop.Value as string;
                }
              }
              eq.Location = br.Position;
              if (addEquip)
              {
                PlaceableElectricalEntity p = new PlaceableElectricalEntity();
                p.Id = eq.Id;
                p.ParentId = eq.ParentId;
                p.Location = eq.Location;
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
          if (pooledEquipment[j].ParentId == pooledEquipment[i].Id)
          {
            PlaceableElectricalEntity equip = pooledEquipment[j];
            equip.ParentDistance =
              Convert.ToInt32(
                Math.Abs(pooledEquipment[j].Location.X - pooledEquipment[i].Location.X)
                  + Math.Abs(pooledEquipment[j].Location.Y - pooledEquipment[i].Location.Y)
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
          if (equipmentList[j].Id == pooledEquipment[i].Id)
          {
            isMatch = true;
            ElectricalEntity.Equipment equip = equipmentList[j];
            if (
              equip.ParentDistance != pooledEquipment[i].ParentDistance
              || equip.Location.X != pooledEquipment[i].Location.X
              || equip.Location.Y != pooledEquipment[i].Location.Y
            )
            {
              equip.ParentDistance = pooledEquipment[i].ParentDistance;
              equip.Location = pooledEquipment[i].Location;
              equipmentList[j] = equip;
              gmepDb.UpdateEquipment(equip);
            }
          }
        }
        if (!isMatch)
        {
          for (int j = 0; j < panelList.Count; j++)
          {
            if (panelList[j].Id == pooledEquipment[i].Id)
            {
              isMatch = true;
              ElectricalEntity.Panel panel = panelList[j];
              if (
                panel.ParentDistance != pooledEquipment[i].ParentDistance
                || panel.Location.X != pooledEquipment[i].Location.X
                || panel.Location.Y != pooledEquipment[i].Location.Y
              )
              {
                panel.ParentDistance = pooledEquipment[i].ParentDistance;
                panel.Location = pooledEquipment[i].Location;
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
            if (transformerList[j].Id == pooledEquipment[i].Id)
            {
              isMatch = true;
              Transformer xfmr = transformerList[j];
              if (
                xfmr.ParentDistance != pooledEquipment[i].ParentDistance
                || xfmr.Location.X != pooledEquipment[i].Location.X
                || xfmr.Location.Y != pooledEquipment[i].Location.Y
              )
              {
                xfmr.ParentDistance = pooledEquipment[i].ParentDistance;
                xfmr.Location = pooledEquipment[i].Location;
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

    private string GetCircuitNo(ListViewItem item)
    {
      string circuitNo = item.SubItems[4].Text;
      string voltage = item.SubItems[6].Text;
      string phase = item.SubItems[7].Text;
      if (phase == "3")
      {
        int c = Int32.Parse(circuitNo);
        return circuitNo + "&" + (c + 2).ToString() + "&" + (c + 4).ToString();
      }
      else if (
        voltage == "208"
        || voltage == "230"
        || voltage == "240"
        || voltage == "460"
        || voltage == "480"
      )
      {
        int c = Int32.Parse(circuitNo);
        return circuitNo + "&" + (c + 2).ToString();
      }
      return circuitNo;
    }

    private void EquipmentListView_MouseDoubleClick(object sender, MouseEventArgs e)
    {
      int numSubItems = equipmentListView.SelectedItems[0].SubItems.Count;
      if (equipmentListView.SelectedItems[0].SubItems[numSubItems - 3].Text == "True")
      {
        return;
      }
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

        string circuitNo = GetCircuitNo(equipmentListView.SelectedItems[0]);
        Point3d? p = PlaceEquipment(
          equipmentListView.SelectedItems[0].SubItems[numSubItems - 2].Text,
          equipmentListView.SelectedItems[0].SubItems[numSubItems - 1].Text,
          equipmentListView.SelectedItems[0].Text,
          EquipmentType.Duplex, // TODO set this based on connection
          equipmentListView.SelectedItems[0].SubItems[numSubItems - 5].Text,
          circuitNo,
          equipmentListView.SelectedItems[0].SubItems[numSubItems - 4].Text
        );
        if (p == null)
        {
          return;
        }
        for (int i = 0; i < equipmentList.Count; i++)
        {
          ElectricalEntity.Equipment equipment = equipmentList[i];
          if (
            equipmentList[i].Id == equipmentListView.SelectedItems[0].SubItems[numSubItems - 2].Text
          )
          {
            equipment.Location = (Point3d)p;
            equipmentList[i] = equipment;
          }
        }
      }
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

        if (equipmentListView.SelectedItems.Count > 0)
        {
          numSubItems = equipmentListView.SelectedItems[0].SubItems.Count;
          foreach (ListViewItem item in equipmentListView.SelectedItems)
          {
            if (item.SubItems[numSubItems - 3].Text != "True")
            {
              string circuitNo = GetCircuitNo(item);
              Point3d? p = PlaceEquipment(
                item.SubItems[numSubItems - 2].Text,
                item.SubItems[numSubItems - 1].Text,
                item.Text,
                EquipmentType.Duplex,
                item.SubItems[numSubItems - 5].Text,
                circuitNo,
                item.SubItems[numSubItems - 4].Text
              );
              if (p == null)
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

        if (equipmentListView.Items.Count > 0)
        {
          numSubItems = equipmentListView.Items[0].SubItems.Count;
          foreach (ListViewItem item in equipmentListView.Items)
          {
            if (item.SubItems[numSubItems - 3].Text != "True")
            {
              string circuitNo = GetCircuitNo(item);
              Point3d? p = PlaceEquipment(
                item.SubItems[numSubItems - 2].Text,
                item.SubItems[numSubItems - 1].Text,
                item.Text,
                EquipmentType.Duplex,
                item.SubItems[numSubItems - 5].Text,
                circuitNo,
                item.SubItems[numSubItems - 4].Text
              );
              if (p == null)
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
        int tableCols = 12;
        tb.SetSize(tableRows, tableCols);
        tb.SetRowHeight(0.25);
        tb.Cells[0, 0].TextString = "ELECTRICAL EQUIPMENT SCHEDULE";
        CellRange range = CellRange.Create(tb, 1, 0, 2, 0);
        tb.MergeCells(range);
        range = CellRange.Create(tb, 1, 1, 2, 1);
        tb.MergeCells(range);
        range = CellRange.Create(tb, 1, 2, 1, 6);
        tb.MergeCells(range);
        range = CellRange.Create(tb, 1, 7, 1, 11);
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
        tb.Columns[9].Width = 0.7;
        tb.Columns[10].Width = 0.7;
        tb.Columns[11].Width = 0.7;
        tb.Cells[1, 7].TextString = "ROUGH-IN";
        tb.Cells[2, 7].TextString = "CONNECTION";
        tb.Cells[2, 8].TextString = "HEIGHT";
        tb.Cells[2, 9].TextString = "CONDUIT";
        tb.Cells[2, 10].TextString = "WIRE";
        tb.Cells[2, 11].TextString = "GROUND";
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
          tb.Cells[row, 0].TextString = equipmentList[i].Name;
          tb.Cells[row, 1].TextString = equipmentList[i].Description.ToUpper();
          tb.Cells[row, 2].TextString = equipmentList[i].Voltage.ToString();
          tb.Cells[row, 3].TextString =
            equipmentList[i].Fla > 0 ? Math.Round(equipmentList[i].Fla, 1).ToString() : "-";
          tb.Cells[row, 4].TextString = equipmentList[i].Hp == "0" ? "-" : equipmentList[i].Hp;
          double mca = (equipmentList[i].Mca);
          if (mca <= 0)
          {
            mca = CADObjectCommands.GetMcaFromFla(equipmentList[i].Fla);
          }
          if (mca <= 0)
          {
            tb.Cells[row, 9].TextString = "V.I.F.";
            tb.Cells[row, 10].TextString = "V.I.F.";
          }
          else if (Int32.TryParse(equipmentList[i].Voltage, out int voltage))
          {
            (string firstLine, string secondLine, string _, string _, string _, string _) =
              CADObjectCommands.GetWireAndConduitSizeText(
                equipmentList[i].Fla,
                mca,
                equipmentList[i].ParentDistance + 10,
                voltage,
                3,
                equipmentList[i].Is3Phase ? 3 : 1
              );
            tb.Cells[row, 5].TextString = mca.ToString();
            tb.Cells[row, 7].TextString = CADObjectCommands.GetConnectionTypeFromFlaVoltage(
              equipmentList[i].Fla,
              voltage,
              equipmentList[i].HasPlug,
              equipmentList[i].Is3Phase
            );
            tb.Cells[row, 8].TextString = equipmentList[i].MountingHeight.ToString() + "\"";
            tb.Cells[row, 9].TextString = firstLine.Substring(0, firstLine.IndexOf(" "));
            tb.Cells[row, 10].TextString = firstLine.Substring(firstLine.IndexOf(";") + 2);
            tb.Cells[row, 11].TextString = secondLine.Replace("PLUS ", "").Replace(" GND.", "");
          }

          tb.Cells[row, 6].TextString = equipmentList[i].Is3Phase ? "3" : "1";
        }
        BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
        btr.AppendEntity(tb);
        tr.AddNewlyCreatedDBObject(tb, true);
        tr.Commit();
      }
    }

    private void CreateEquipmentSchedule_Click(object sender, EventArgs e)
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
      var currentView = ed.GetCurrentView();
      using (
        DocumentLock docLock =
          Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
      )
      {
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          var promptOptions = new PromptPointOptions("\nSelect upper left point:");
          var promptResult = ed.GetPoint(promptOptions);
          currentView = ed.GetCurrentView();
          if (promptResult.Status == PromptStatus.OK)
            startingPoint = promptResult.Value;
          else
          {
            return;
          }
        }
        CreateEquipmentSchedule(doc, db, ed, startingPoint);
      }
      ed.SetCurrentView(currentView);
    }
  }
}
