﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using DocumentFormat.OpenXml.Drawing.Charts;
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
    public double aicRating;

    public Service(
      string id,
      string name,
      string meterConfig,
      int amp,
      string voltage,
      double aicRating
    )
    {
      this.id = id;
      this.name = name;
      isMultiMeter = meterConfig == "MULTIMETER";
      this.amp = amp;
      this.voltage = voltage;
      this.aicRating = aicRating;
    }
  }

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
      projectId = gmepDb.GetProjectId(CADObjectCommands.GetProjectNoFromFileName());
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
        item.SubItems.Add(equipment.circuit.ToString());
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
        item.SubItems.Add(equipment.hidden.ToString());
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
        equipmentListView.Columns.Add("Circuit", -2, HorizontalAlignment.Left);
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
        item.SubItems.Add(panel.hidden.ToString());
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
        item.SubItems.Add(xfmr.hidden.ToString());
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
          if (!equipmentList[i].hidden)
          {
            int parentDistance = -1;
            foreach (Panel p in panelList)
            {
              if (p.id == equipmentList[i].parentId)
              {
                if (p.hidden)
                {
                  parentDistance = equipmentList[i].parentDistance;
                }
              }
            }
            foreach (Transformer t in transformerList)
            {
              if (t.id == equipmentList[i].parentId)
              {
                if (t.hidden)
                {
                  parentDistance = equipmentList[i].parentDistance;
                }
              }
            }
            eq.parentDistance = parentDistance;
          }
          else
          {
            eq.parentDistance = equipmentList[i].parentDistance;
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
          if (!panelList[i].hidden)
          {
            int parentDistance = -1;
            foreach (Panel p in panelList)
            {
              if (p.id == panelList[i].parentId)
              {
                if (p.hidden)
                {
                  parentDistance = panelList[i].parentDistance;
                }
              }
            }
            foreach (Transformer t in transformerList)
            {
              if (t.id == panelList[i].parentId)
              {
                if (t.hidden)
                {
                  parentDistance = panelList[i].parentDistance;
                }
              }
            }
            panel.parentDistance = parentDistance;
          }
          else
          {
            panel.parentDistance = panelList[i].parentDistance;
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
          if (!transformerList[i].hidden)
          {
            int parentDistance = -1;
            foreach (Panel p in panelList)
            {
              if (p.id == transformerList[i].parentId)
              {
                if (p.hidden)
                {
                  parentDistance = transformerList[i].parentDistance;
                }
              }
            }

            xfmr.parentDistance = parentDistance;
          }
          else
          {
            xfmr.parentDistance = transformerList[i].parentDistance;
          }
          transformerList[i] = xfmr;
          gmepDb.UpdateTransformer(xfmr);
        }
      }
    }

    private Point3d? PlaceEquipment(
      string equipId,
      string parentId,
      string equipNo,
      EquipmentType equipType,
      string circuitNo = ""
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
      PromptPointOptions promptOptions = new PromptPointOptions(
        "\nSelect point for " + equipNo + ": "
      );
      PromptPointResult promptResult = ed.GetPoint(promptOptions);
      if (promptResult.Status != PromptStatus.OK)
        return null;
      Point3d firstClickPoint = promptResult.Value;
      // insert block here
      string blockName = $"A$C3D4728D6";
      double rotation = 0;
      switch (equipType)
      {
        case EquipmentType.Duplex:
          blockName = $"GMEP DUPLEX";
          break;
        case EquipmentType.Panel:
          blockName = $"A$C26441056";
          break;
        // TODO check for remaining connection types
      }
      using (Transaction acTrans = db.TransactionManager.StartTransaction())
      {
        BlockTable acBlkTbl = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord acBlkTblRec =
          acTrans.GetObject(acBlkTbl[blockName], OpenMode.ForRead) as BlockTableRecord;

        using (BlockReference acBlkRef = new BlockReference(Point3d.Origin, acBlkTblRec.ObjectId))
        {
          RotateJig blockJig = new RotateJig(acBlkRef);
          PromptResult blockPromptResult = ed.Drag(blockJig);
          double scaleFactor = 0.25;
          if (equipType != EquipmentType.Panel && equipType != EquipmentType.Transformer)
          {
            scaleFactor = scale;
          }
          if (promptResult.Status == PromptStatus.OK)
          {
            BlockTableRecord currentSpace =
              acTrans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
            currentSpace.AppendEntity(acBlkRef);
            acTrans.AddNewlyCreatedDBObject(acBlkRef, true);
            acBlkRef.ScaleFactors = new Scale3d(0.25 / scaleFactor);
            acBlkRef.Layer = "E-SYM1";
            acTrans.Commit();
            rotation = acBlkRef.Rotation;
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
            circuitOffsetY = -9;
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
            circuitOffsetX = 9;
            circuitOffsetY = -4.5;
            break;
          default:
            circuitOffsetY = -9;
            circuitOffsetX = 4.5;
            break;
        }
        circuitOffsetY = circuitOffsetY * 0.25 / scale;
        circuitOffsetX = circuitOffsetX * 0.25 / scale;
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
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
          tr.Commit();
        }
      }

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
          ObjectId blockId = bt["EQUIPMENT LOCATOR"];
          using (BlockReference acBlkRef = new BlockReference(firstClickPoint, blockId))
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
              if (prop.PropertyName == "gmep_equip_no")
              {
                prop.Value = equipNo;
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
                if (prop.PropertyName == "gmep_equip_locator" && prop.Value as string == "true")
                {
                  addEquip = true;
                }
                if (prop.PropertyName == "gmep_equip_id" && prop.Value as string != "0")
                {
                  eq.id = prop.Value as string;
                }
                if (prop.PropertyName == "gmep_equip_parent_id" && prop.Value as string != "0")
                {
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
          circuitNo
        );
        if (p == null)
        {
          return;
        }
        for (int i = 0; i < equipmentList.Count; i++)
        {
          Equipment equipment = equipmentList[i];
          if (
            equipmentList[i].id == equipmentListView.SelectedItems[0].SubItems[numSubItems - 2].Text
          )
          {
            equipment.loc = (Point3d)p;
            equipmentList[i] = equipment;
          }
        }
      }
      CalculateDistances();
    }

    private void PanelListView_MouseDoubleClick(object sender, MouseEventArgs e)
    {
      int numSubItems = panelListView.SelectedItems[0].SubItems.Count;
      if (panelListView.SelectedItems[0].SubItems[numSubItems - 3].Text == "True")
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
        Point3d? p = PlaceEquipment(
          panelListView.SelectedItems[0].SubItems[numSubItems - 2].Text,
          panelListView.SelectedItems[0].SubItems[numSubItems - 1].Text,
          panelListView.SelectedItems[0].Text,
          EquipmentType.Panel
        );
        if (p == null)
        {
          return;
        }
      }
      CreatePanelListView(true);
      CalculateDistances();
    }

    private void TransformerListView_MouseDoubleClick(object sender, MouseEventArgs e)
    {
      int numSubItems = transformerListView.SelectedItems[0].SubItems.Count;
      if (transformerListView.SelectedItems[0].SubItems[numSubItems - 3].Text == "True")
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
        Point3d? p = PlaceEquipment(
          transformerListView.SelectedItems[0].SubItems[numSubItems - 2].Text,
          transformerListView.SelectedItems[0].SubItems[numSubItems - 1].Text,
          transformerListView.SelectedItems[0].Text,
          EquipmentType.Transformer
        );
        if (p == null)
        {
          return;
        }
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
          if (item.SubItems[numSubItems - 3].Text != "True")
          {
            Point3d? p = PlaceEquipment(
              item.SubItems[numSubItems - 2].Text,
              item.SubItems[numSubItems - 1].Text,
              item.Text,
              EquipmentType.Panel
            );
            if (p == null)
            {
              brk = true;
              break;
            }
            panelLocs[item.SubItems[numSubItems - 2].Text] = (Point3d)p;
          }
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
              if (item.SubItems[numSubItems - 3].Text != "True")
              {
                Point3d? p = PlaceEquipment(
                  item.SubItems[numSubItems - 2].Text,
                  item.SubItems[numSubItems - 1].Text,
                  item.Text,
                  EquipmentType.Transformer
                );
                if (p == null)
                {
                  return;
                }
                if (p == null)
                {
                  brk = true;
                  break;
                }
              }
            }
          }
          if (!brk && equipmentListView.SelectedItems.Count > 0)
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
                  circuitNo
                );
                if (p == null)
                {
                  break;
                }
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
          if (item.SubItems[numSubItems - 3].Text != "True")
          {
            Point3d? p = PlaceEquipment(
              item.SubItems[numSubItems - 2].Text,
              item.SubItems[numSubItems - 1].Text,
              item.Text,
              EquipmentType.Panel
            );
            if (p == null)
            {
              brk = true;
              break;
            }
          }
        }
        if (!brk && (transformerListView.Items.Count > 0 || equipmentListView.Items.Count > 0))
        {
          if (transformerListView.Items.Count > 0)
          {
            numSubItems = transformerListView.Items[0].SubItems.Count;
            foreach (ListViewItem item in transformerListView.Items)
            {
              if (item.SubItems[numSubItems - 3].Text != "True")
              {
                Point3d? p = PlaceEquipment(
                  item.SubItems[numSubItems - 2].Text,
                  item.SubItems[numSubItems - 1].Text,
                  item.Text,
                  EquipmentType.Transformer
                );
                if (p == null)
                {
                  brk = true;
                  break;
                }
              }
            }
          }
          if (!brk && equipmentListView.Items.Count > 0)
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
                  circuitNo
                );
                if (p == null)
                {
                  break;
                }
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
          childPanel.parentAicRating = panel.aicRating;
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
          childXfmr.kva = t.kva;
          childXfmr.parentAicRating = panel.aicRating;
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
          childPanel.mainBreakerSize = p.busSize;
          childPanel.voltageSpec = p.voltage;
          if (p.voltage.Contains("3"))
          {
            childPanel.is3Phase = true;
          }
          childPanel.parentAicRating = transformer.aicRating;
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
          if (sf.amp >= 1200 && sf.voltageSpec.Contains("480"))
          {
            p.hasGfp = true;
          }
          p.distributionBreakerSize = sf.amp;
          p.voltageSpec = sf.voltageSpec;
          p.is3Phase = false;
          p.aicRating = sf.aicRating;
          if (sf.voltageSpec.Contains("3"))
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
        sf.aicRating = service.aicRating;
        MakeSingleLineNodeTreeFromService(sf);
        singleLine.children.Add(sf);
      }
      return singleLine;
    }

    private void MakeSingleLineButton_Click(object sender, EventArgs e)
    {
      SingleLine singleLineNodeTree;
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
        singleLineNodeTree = MakeSingleLineNodeTree();
        singleLineNodeTree.AggregateWidths();
        singleLineNodeTree.SetChildStartingPoints(startingPoint);
        singleLineNodeTree.Make();
      }
      singleLineNodeTree.SaveAicRatings();
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
          tb.Cells[row, 0].TextString = equipmentList[i].name;
          tb.Cells[row, 1].TextString = equipmentList[i].description.ToUpper();
          tb.Cells[row, 2].TextString = equipmentList[i].voltage.ToString();
          tb.Cells[row, 3].TextString =
            equipmentList[i].fla > 0 ? Math.Round(equipmentList[i].fla, 1).ToString() : "-";
          tb.Cells[row, 4].TextString = equipmentList[i].hp == "0" ? "-" : equipmentList[i].hp;
          double mca = (equipmentList[i].mca);
          if (mca <= 0)
          {
            mca = CADObjectCommands.GetMcaFromFla(equipmentList[i].fla);
          }
          if (mca <= 0)
          {
            tb.Cells[row, 9].TextString = "V.I.F.";
            tb.Cells[row, 10].TextString = "V.I.F.";
          }
          else
          {
            (string firstLine, string secondLine, string _, string _, string _, string _) =
              CADObjectCommands.GetWireAndConduitSizeText(
                equipmentList[i].fla,
                mca,
                equipmentList[i].parentDistance + 10,
                equipmentList[i].voltage,
                3,
                equipmentList[i].is3Phase ? 3 : 1
              );
            tb.Cells[row, 5].TextString = mca.ToString();
            tb.Cells[row, 7].TextString = CADObjectCommands.GetConnectionTypeFromFlaVoltage(
              equipmentList[i].fla,
              equipmentList[i].voltage,
              equipmentList[i].hasPlug,
              equipmentList[i].is3Phase
            );
            tb.Cells[row, 8].TextString = equipmentList[i].mountingHeight.ToString() + "\"";
            tb.Cells[row, 9].TextString = firstLine.Substring(0, firstLine.IndexOf(" "));
            tb.Cells[row, 10].TextString = firstLine.Substring(firstLine.IndexOf(";") + 2);
            tb.Cells[row, 11].TextString = secondLine.Replace("PLUS ", "").Replace(" GND.", "");
          }

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
