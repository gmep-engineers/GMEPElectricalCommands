using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using DocumentFormat.OpenXml.Drawing.Charts;
using ElectricalCommands.ElectricalEntity;
using GMEPElectricalCommands.GmepDatabase;

namespace ElectricalCommands.SingleLine
{
  enum GroupType
  {
    DistributionSection,
    MultimeterSection,
    PullSection,
    MainBreakerSection,
    MainMeterSection,
    MainMeterAndBreakerSection,
  }

  public partial class SingleLineDialogWindow : Form
  {
    private string projectId;
    private List<ElectricalEntity.Service> serviceList;
    private List<ElectricalEntity.Meter> meterList;
    private List<ElectricalEntity.MainBreaker> mainBreakerList;
    private List<ElectricalEntity.DistributionBus> distributionBusList;
    private List<ElectricalEntity.DistributionBreaker> distributionBreakerList;
    private List<ElectricalEntity.Panel> panelList;
    private List<ElectricalEntity.PanelBreaker> panelBreakerList;
    private List<ElectricalEntity.Disconnect> disconnectList;
    private List<ElectricalEntity.Transformer> transformerList;
    private List<ElectricalEntity.Equipment> equipmentList;
    private List<ElectricalEntity.NodeLink> nodeLinkList;
    private List<ElectricalEntity.GroupNode> groupList;
    private Dictionary<string, List<string>> groupDict;
    private List<ObjectId> dynamicBlockIds;

    private List<PlaceableElectricalEntity> placeables;
    public GmepDatabase gmepDb;

    public SingleLineDialogWindow(SingleLineCommands singleLineCommands)
    {
      InitializeComponent();
    }

    public void InitializeModal()
    {
      gmepDb = new GmepDatabase();
      projectId = gmepDb.GetProjectId(CADObjectCommands.GetProjectNoFromFileName());
      serviceList = gmepDb.GetServices(projectId);
      meterList = gmepDb.GetMeters(projectId);
      mainBreakerList = gmepDb.GetMainBreakers(projectId);
      distributionBusList = gmepDb.GetDistributionBuses(projectId);
      distributionBreakerList = gmepDb.GetDistributionBreakers(projectId);
      panelList = gmepDb.GetPanels(projectId);
      panelBreakerList = gmepDb.GetPanelBreakers(projectId);
      disconnectList = gmepDb.GetDisconnects(projectId);
      transformerList = gmepDb.GetTransformers(projectId);
      equipmentList = gmepDb.GetEquipment(projectId, true);
      nodeLinkList = gmepDb.GetNodeLinks(projectId);
      groupList = gmepDb.GetGroupNodes(projectId);

      MakeGroupDict();

      SingleLineTreeView.BeginUpdate();
      PopulateTreeView();
      SingleLineTreeView.ExpandAll();
      SingleLineTreeView.EndUpdate();
      if (serviceList.Count > 0)
      {
        SetInfoBoxText(serviceList[0]);
      }
      placeables = new List<PlaceableElectricalEntity>();
      placeables.AddRange(serviceList);
      placeables.AddRange(distributionBusList);
      placeables.AddRange(panelList);
      placeables.AddRange(disconnectList);
      placeables.AddRange(transformerList);
      dynamicBlockIds = new List<ObjectId>();
      SetDynamicBlockIds();
      placeables.ForEach(placeable => placeable.SetBlockId(dynamicBlockIds));
      foreach (TreeNode node in SingleLineTreeView.Nodes)
      {
        if (node != null)
        {
          ContextMenuStrip docMenu = new ContextMenuStrip();
          ToolStripMenuItem placeLabel = new ToolStripMenuItem();
          placeLabel.Text = "Place";
          placeLabel.Click += TreeViewPlaceSelected_Click;
          docMenu.Items.Add(placeLabel);
          node.ContextMenuStrip = docMenu;
          SetTreeContextMenu(node);
        }
      }
    }

    private void SetTreeContextMenu(TreeNode parentNode)
    {
      foreach (TreeNode node in parentNode.Nodes)
      {
        if (node != null)
        {
          ContextMenuStrip docMenu = new ContextMenuStrip();
          ToolStripMenuItem placeLabel = new ToolStripMenuItem();
          placeLabel.Text = "Place";
          placeLabel.Click += TreeViewPlaceSelected_Click;
          docMenu.Items.Add(placeLabel);
          node.ContextMenuStrip = docMenu;
          SetTreeContextMenu(node);
        }
      }
    }

    public void SetDynamicBlockIds()
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
      dynamicBlockIds.Clear();
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
              dynamicBlockIds.Add(id);
            }
          }
          catch { }
        }
      }
    }

    public void MakeGroupDict()
    {
      groupDict = new Dictionary<string, List<string>>();
      groupList.ForEach(group => groupDict.Add(group.Id, new List<string>()));
      groupDict.Add(String.Empty, new List<string>());
      serviceList.ForEach(service => groupDict[GetGroupAssociation(service)].Add(service.Id));
      meterList.ForEach(meter => groupDict[GetGroupAssociation(meter)].Add(meter.Id));
      mainBreakerList.ForEach(mainBreaker =>
        groupDict[GetGroupAssociation(mainBreaker)].Add(mainBreaker.Id)
      );
      distributionBusList.ForEach(distributionBus =>
        groupDict[GetGroupAssociation(distributionBus)].Add(distributionBus.Id)
      );
      distributionBreakerList.ForEach(distributionBreaker =>
        groupDict[GetGroupAssociation(distributionBreaker)].Add(distributionBreaker.Id)
      );
    }

    public double AggregateEntityWidth(TreeNode parentNode, NodeType parentType)
    {
      double width = 0;
      if (parentType == NodeType.Undefined)
      {
        width = 2;
      }
      foreach (TreeNode childNode in parentNode.Nodes)
      {
        ElectricalEntity.ElectricalEntity childEntity = (ElectricalEntity.ElectricalEntity)
          childNode.Tag;
        switch (childEntity.NodeType)
        {
          case NodeType.Meter:
            width += AggregateEntityWidth(childNode, childEntity.NodeType);
            break;
          case NodeType.DistributionBreaker:
            width += AggregateEntityWidth(childNode, childEntity.NodeType);
            break;
          case NodeType.Panel:
            width +=
              AggregateEntityWidth(childNode, childEntity.NodeType)
              + (parentType == NodeType.Panel ? 2 : 0);
            break;
          case NodeType.Disconnect:
            width +=
              AggregateEntityWidth(childNode, childEntity.NodeType)
              + (parentType == NodeType.Panel ? 2 : 0);
            break;
          case NodeType.Transformer:
            width +=
              AggregateEntityWidth(childNode, childEntity.NodeType)
              + (parentType == NodeType.Panel ? 2 : 0);
            break;
          case NodeType.PanelBreaker:
            width += 2 + AggregateEntityWidth(childNode, childEntity.NodeType);
            break;
        }
      }
      return width;
    }

    public double AggregateGroupWidth(string groupId)
    {
      double width = 0;
      foreach (string childId in groupDict[groupId])
      {
        TreeNode[] nodes = SingleLineTreeView.Nodes.Find(childId, true);
        if (nodes == null || nodes.Length == 0)
        {
          return width;
        }
        TreeNode childNode = nodes[0];
        ElectricalEntity.ElectricalEntity parent = (ElectricalEntity.ElectricalEntity)
          childNode.Parent.Tag;
        if (parent.NodeType == NodeType.DistributionBus)
        {
          width += AggregateEntityWidth(childNode, NodeType.Undefined);
        }
      }
      return width;
    }

    public double AggregateDistributionBusWidth(string distributionBusId)
    {
      double width = 2;
      TreeNode[] nodes = SingleLineTreeView.Nodes.Find(distributionBusId, true);
      if (nodes == null || nodes.Length == 0)
      {
        return width;
      }
      TreeNode thisNode = nodes[0];
      foreach (TreeNode childNode in thisNode.Nodes)
      {
        ElectricalEntity.ElectricalEntity childEntity = (ElectricalEntity.ElectricalEntity)
          childNode.Tag;
        width += AggregateEntityWidth(childNode, childEntity.NodeType);
      }
      return width;
    }

    public string GetGroupAssociation(ElectricalEntity.ElectricalEntity entity)
    {
      Point entityPoint = entity.NodePosition;
      foreach (ElectricalEntity.GroupNode group in groupList)
      {
        Point groupPoint = group.NodePosition;
        if (
          groupPoint.X + group.Width > entityPoint.X
          && entityPoint.X > groupPoint.X
          && groupPoint.Y + group.Height > entityPoint.Y
          && entityPoint.Y > groupPoint.Y
        )
        {
          return group.Id;
        }
      }
      return String.Empty;
    }

    public void SetTreeNodeColor(TreeNode node, PlaceableElectricalEntity entity)
    {
      if (entity.NodeType == NodeType.Equipment)
      {
        return;
      }
      if (entity.Location.X == 0 && entity.Location.Y == 0)
      {
        node.ForeColor = Color.White;
        node.BackColor = Color.Crimson;
      }
      else
      {
        node.ForeColor = Color.White;
        node.BackColor = Color.DarkCyan;
      }
    }

    public void InheritElectricalAttributes(
      ElectricalEntity.ElectricalEntity parent,
      ElectricalEntity.ElectricalEntity child
    )
    {
      child.LineVoltage = parent.LineVoltage;
      child.Phase = parent.Phase;
      child.AicRating = parent.AicRating;
    }

    public void PopulateTreeView()
    {
      foreach (ElectricalEntity.Service service in serviceList)
      {
        TreeNode serviceNode = SingleLineTreeView.Nodes.Add(
          service.Id,
          service.Name.Replace("\u0081", "\u03A6")
        );
        serviceNode.Tag = service;
        SetTreeNodeColor(serviceNode, service);
        PopulateFromService(serviceNode, service);
      }
    }

    public void PopulateFromService(TreeNode node, ElectricalEntity.Service service)
    {
      foreach (ElectricalEntity.Meter meter in meterList)
      {
        if (VerifyNodeLink(service.NodeId, meter.NodeId))
        {
          InheritElectricalAttributes(service, meter);
          TreeNode meterNode = node.Nodes.Add(meter.Id, meter.Name);
          meterNode.Tag = meter;
          PopulateFromMainMeter(meterNode, meter);
        }
      }
      foreach (ElectricalEntity.MainBreaker mainBreaker in mainBreakerList)
      {
        if (VerifyNodeLink(service.NodeId, mainBreaker.NodeId))
        {
          InheritElectricalAttributes(service, mainBreaker);
          TreeNode mainBreakerNode = node.Nodes.Add(mainBreaker.Id, mainBreaker.Name);
          mainBreakerNode.Tag = mainBreaker;
          PopulateFromMainBreaker(mainBreakerNode, mainBreaker);
        }
      }
    }

    public void PopulateFromMainMeter(TreeNode node, ElectricalEntity.Meter meter)
    {
      foreach (ElectricalEntity.MainBreaker mainBreaker in mainBreakerList)
      {
        if (VerifyNodeLink(meter.NodeId, mainBreaker.NodeId))
        {
          InheritElectricalAttributes(meter, mainBreaker);
          TreeNode mainBreakerNode = node.Nodes.Add(mainBreaker.Id, mainBreaker.Name);
          mainBreakerNode.Tag = mainBreaker;
          PopulateFromMainBreaker(mainBreakerNode, mainBreaker);
        }
      }
      foreach (ElectricalEntity.DistributionBus distributionBus in distributionBusList)
      {
        if (VerifyNodeLink(meter.NodeId, distributionBus.NodeId))
        {
          InheritElectricalAttributes(meter, distributionBus);
          TreeNode distributionBusNode = node.Nodes.Add(distributionBus.Id, distributionBus.Name);
          distributionBusNode.Tag = distributionBus;
          SetTreeNodeColor(distributionBusNode, distributionBus);
          PopulateFromDistributionBus(distributionBusNode, distributionBus);
        }
      }
    }

    public void PopulateFromMainBreaker(TreeNode node, ElectricalEntity.MainBreaker mainBreaker)
    {
      foreach (ElectricalEntity.DistributionBus distributionBus in distributionBusList)
      {
        if (VerifyNodeLink(mainBreaker.NodeId, distributionBus.NodeId))
        {
          InheritElectricalAttributes(mainBreaker, distributionBus);
          TreeNode distributionBusNode = node.Nodes.Add(distributionBus.Id, distributionBus.Name);
          distributionBusNode.Tag = distributionBus;
          SetTreeNodeColor(distributionBusNode, distributionBus);
          PopulateFromDistributionBus(distributionBusNode, distributionBus);
        }
      }
    }

    public void PopulateFromDistributionBus(
      TreeNode node,
      ElectricalEntity.DistributionBus distributionBus
    )
    {
      foreach (ElectricalEntity.Meter meter in meterList)
      {
        if (VerifyNodeLink(distributionBus.NodeId, meter.NodeId))
        {
          InheritElectricalAttributes(distributionBus, meter);
          TreeNode distributionMeterNode = node.Nodes.Add(meter.Id, meter.Name);
          distributionMeterNode.Tag = meter;
          PopulateFromDistributionMeter(distributionMeterNode, meter);
        }
      }
      foreach (ElectricalEntity.DistributionBreaker distributionBreaker in distributionBreakerList)
      {
        if (VerifyNodeLink(distributionBus.NodeId, distributionBreaker.NodeId))
        {
          InheritElectricalAttributes(distributionBus, distributionBreaker);
          TreeNode distributionBreakerNode = node.Nodes.Add(
            distributionBreaker.Id,
            distributionBreaker.Name
          );
          distributionBreakerNode.Tag = distributionBreaker;
          PopulateFromDistributionBreaker(distributionBreakerNode, distributionBreaker);
        }
      }
    }

    public void PopulateFromDistributionMeter(TreeNode node, ElectricalEntity.Meter meter)
    {
      foreach (ElectricalEntity.DistributionBreaker distributionBreaker in distributionBreakerList)
      {
        if (VerifyNodeLink(meter.NodeId, distributionBreaker.NodeId))
        {
          InheritElectricalAttributes(meter, distributionBreaker);
          TreeNode distributionBreakerNode = node.Nodes.Add(
            distributionBreaker.Id,
            distributionBreaker.Name
          );
          distributionBreakerNode.Tag = distributionBreaker;
          PopulateFromDistributionBreaker(distributionBreakerNode, distributionBreaker);
        }
      }
    }

    public void PopulateFromDistributionBreaker(
      TreeNode node,
      ElectricalEntity.DistributionBreaker distributionBreaker
    )
    {
      foreach (ElectricalEntity.Panel panel in panelList)
      {
        if (VerifyNodeLink(distributionBreaker.NodeId, panel.NodeId))
        {
          InheritElectricalAttributes(distributionBreaker, panel);
          TreeNode panelNode = node.Nodes.Add(panel.Id, "Panel " + panel.Name);
          panelNode.Tag = panel;
          SetTreeNodeColor(panelNode, panel);
          PopulateFromPanel(panelNode, panel);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(distributionBreaker.NodeId, disconnect.NodeId))
        {
          InheritElectricalAttributes(distributionBreaker, disconnect);
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          disconnectNode.Tag = disconnect;
          SetTreeNodeColor(disconnectNode, disconnect);
          PopulateFromDisconnect(disconnectNode, disconnect);
        }
      }
      foreach (ElectricalEntity.Transformer transformer in transformerList)
      {
        if (VerifyNodeLink(distributionBreaker.NodeId, transformer.NodeId))
        {
          TreeNode transformerNode = node.Nodes.Add(
            transformer.Id,
            transformer.Name.Replace("\u0081", "\u03A6")
          );
          transformerNode.Tag = transformer;
          SetTreeNodeColor(transformerNode, transformer);
          PopulateFromTransformer(transformerNode, transformer);
        }
      }
    }

    public void PopulateFromPanel(TreeNode node, ElectricalEntity.Panel panel)
    {
      foreach (ElectricalEntity.PanelBreaker panelBreaker in panelBreakerList)
      {
        if (VerifyNodeLink(panel.NodeId, panelBreaker.NodeId))
        {
          InheritElectricalAttributes(panel, panelBreaker);
          TreeNode panelBreakerNode = node.Nodes.Add(panelBreaker.Id, panelBreaker.Name);
          panelBreakerNode.Tag = panelBreaker;
          PopulateFromPanelBreaker(panelBreakerNode, panelBreaker);
        }
      }
      foreach (ElectricalEntity.Panel childPanel in panelList)
      {
        if (VerifyNodeLink(panel.NodeId, childPanel.NodeId))
        {
          InheritElectricalAttributes(panel, childPanel);
          TreeNode childPanelNode = node.Nodes.Add(childPanel.Id, "Panel " + childPanel.Name);
          childPanelNode.Tag = childPanel;
          SetTreeNodeColor(childPanelNode, childPanel);
          PopulateFromPanel(childPanelNode, childPanel);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(panel.NodeId, disconnect.NodeId))
        {
          InheritElectricalAttributes(panel, disconnect);
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          disconnectNode.Tag = disconnect;
          SetTreeNodeColor(disconnectNode, disconnect);
          PopulateFromDisconnect(disconnectNode, disconnect);
        }
      }
      foreach (ElectricalEntity.Transformer transformer in transformerList)
      {
        if (VerifyNodeLink(panel.NodeId, transformer.NodeId))
        {
          TreeNode transformerNode = node.Nodes.Add(
            transformer.Id,
            transformer.Name.Replace("\u0081", "\u03A6")
          );
          transformerNode.Tag = transformer;
          SetTreeNodeColor(transformerNode, transformer);
          PopulateFromTransformer(transformerNode, transformer);
        }
      }
    }

    public void PopulateFromPanelBreaker(TreeNode node, ElectricalEntity.PanelBreaker panelBreaker)
    {
      foreach (ElectricalEntity.Panel panel in panelList)
      {
        if (VerifyNodeLink(panelBreaker.NodeId, panel.NodeId))
        {
          InheritElectricalAttributes(panelBreaker, panel);
          TreeNode panelNode = node.Nodes.Add(panel.Id, "Panel " + panel.Name);
          panelNode.Tag = panel;
          SetTreeNodeColor(panelNode, panel);
          PopulateFromPanel(panelNode, panel);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(panelBreaker.NodeId, disconnect.NodeId))
        {
          InheritElectricalAttributes(panelBreaker, disconnect);
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          disconnectNode.Tag = disconnect;
          SetTreeNodeColor(disconnectNode, disconnect);
          PopulateFromDisconnect(disconnectNode, disconnect);
        }
      }
      foreach (ElectricalEntity.Transformer transformer in transformerList)
      {
        if (VerifyNodeLink(panelBreaker.NodeId, transformer.NodeId))
        {
          TreeNode transformerNode = node.Nodes.Add(
            transformer.Id,
            transformer.Name.Replace("\u0081", "\u03A6")
          );
          transformerNode.Tag = transformer;
          SetTreeNodeColor(transformerNode, transformer);
          PopulateFromTransformer(transformerNode, transformer);
        }
      }
    }

    public void PopulateFromDisconnect(TreeNode node, ElectricalEntity.Disconnect disconnect)
    {
      foreach (ElectricalEntity.Panel panel in panelList)
      {
        if (VerifyNodeLink(disconnect.NodeId, panel.NodeId))
        {
          InheritElectricalAttributes(disconnect, panel);
          TreeNode panelNode = node.Nodes.Add(panel.Id, "Panel " + panel.Name);
          panelNode.Tag = panel;
          SetTreeNodeColor(panelNode, panel);
          PopulateFromPanel(panelNode, panel);
        }
      }
      foreach (ElectricalEntity.Disconnect childDisconnect in disconnectList)
      {
        if (VerifyNodeLink(disconnect.NodeId, childDisconnect.NodeId))
        {
          InheritElectricalAttributes(disconnect, childDisconnect);
          TreeNode childDisconnectNode = node.Nodes.Add(childDisconnect.Id, childDisconnect.Name);
          childDisconnectNode.Tag = childDisconnect;
          SetTreeNodeColor(childDisconnectNode, childDisconnect);
          PopulateFromDisconnect(childDisconnectNode, childDisconnect);
        }
      }
      foreach (ElectricalEntity.Transformer transformer in transformerList)
      {
        if (VerifyNodeLink(disconnect.NodeId, transformer.NodeId))
        {
          TreeNode transformerNode = node.Nodes.Add(
            transformer.Id,
            transformer.Name.Replace("\u0081", "\u03A6")
          );
          transformerNode.Tag = transformer;
          SetTreeNodeColor(transformerNode, transformer);
          PopulateFromTransformer(transformerNode, transformer);
        }
      }
      foreach (ElectricalEntity.Equipment equipment in equipmentList)
      {
        if (VerifyNodeLink(disconnect.NodeId, equipment.NodeId))
        {
          InheritElectricalAttributes(disconnect, equipment);
          TreeNode equipmentNode = node.Nodes.Add(equipment.Id, equipment.Name);
          equipmentNode.Tag = equipment;
          SetTreeNodeColor(equipmentNode, equipment);
        }
      }
    }

    public void PopulateFromTransformer(TreeNode node, ElectricalEntity.Transformer transformer)
    {
      foreach (ElectricalEntity.Panel panel in panelList)
      {
        if (VerifyNodeLink(transformer.NodeId, panel.NodeId))
        {
          panel.LineVoltage = transformer.OutputLineVoltage;
          panel.Phase = transformer.Phase; // todo: account for phase change
          TreeNode panelNode = node.Nodes.Add(panel.Id, "Panel " + panel.Name);
          panelNode.Tag = panel;
          SetTreeNodeColor(panelNode, panel);
          PopulateFromPanel(panelNode, panel);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(transformer.NodeId, disconnect.NodeId))
        {
          disconnect.LineVoltage = transformer.OutputLineVoltage;
          disconnect.Phase = transformer.Phase; // todo: account for phase change
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          disconnectNode.Tag = disconnect;
          SetTreeNodeColor(disconnectNode, disconnect);
          PopulateFromDisconnect(disconnectNode, disconnect);
        }
      }
    }

    public bool VerifyNodeLink(string outputConnectorNodeId, string inputConnectorNodeId)
    {
      foreach (ElectricalEntity.NodeLink link in nodeLinkList)
      {
        if (
          link.OutputConnectorNodeId == outputConnectorNodeId
          && link.InputConnectorNodeId == inputConnectorNodeId
        )
        {
          return true;
        }
      }
      return false;
    }

    private string GetParentName(string parentId)
    {
      string parentName = string.Empty;
      PlaceableElectricalEntity parent = placeables.FirstOrDefault(p => parentId == p.Id);
      if (parent != null)
      {
        parentName = parent.Name;
        if (parent.NodeType == NodeType.Panel)
        {
          parentName = "Panel " + parentName;
        }
      }
      return parentName;
    }

    private void SetInfoBoxText(ElectricalEntity.ElectricalEntity entity)
    {
      InfoTextBox.Clear();
      InfoGroupBox.Text = entity.Name.Replace("\u0081", "\u03A6");
      InfoTextBox.AppendText("--------------------General---------------------");
      InfoTextBox.AppendText(Environment.NewLine);
      InfoTextBox.AppendText($"ID:       {entity.Id}");
      InfoTextBox.AppendText(Environment.NewLine);
      InfoTextBox.AppendText($"Status:   {entity.Status}");
      InfoTextBox.AppendText(Environment.NewLine);
      InfoTextBox.AppendText($"AIC:      {entity.AicRating} AIC");
      InfoTextBox.AppendText(Environment.NewLine);
      switch (entity.NodeType)
      {
        case NodeType.Service:
          ElectricalEntity.Service service = (ElectricalEntity.Service)entity;
          InfoTextBox.AppendText($"Location: {GetLocationString(service)}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText("--------------------Service---------------------");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Amp Rating: {service.AmpRating}A");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Voltage:    {service.Voltage.Replace(" ", "V-") + "\u03A6"}");
          break;
        case NodeType.MainBreaker:
          ElectricalEntity.MainBreaker mainBreaker = (ElectricalEntity.MainBreaker)entity;
          InfoTextBox.AppendText("--------------------Breaker---------------------");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Amp Rating: {mainBreaker.AmpRating}A");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Poles:      {mainBreaker.NumPoles}P");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText("----------------Protection Types----------------");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText(
            "Ground Fault: " + (mainBreaker.HasGroundFaultProtection ? "Yes" : "No")
          );
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText(
            "Surge:        " + (mainBreaker.HasSurgeProtection ? "Yes" : "No")
          );
          break;
        case NodeType.DistributionBus:
          ElectricalEntity.DistributionBus distributionBus =
            (ElectricalEntity.DistributionBus)entity;
          InfoTextBox.AppendText($"Location: {GetLocationString(distributionBus)}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText("----------------------Bus-----------------------");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Amp Rating: {distributionBus.AmpRating}A");
          break;
        case NodeType.DistributionBreaker:
          ElectricalEntity.DistributionBreaker distributionBreaker =
            (ElectricalEntity.DistributionBreaker)entity;
          InfoTextBox.AppendText("--------------------Breaker---------------------");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Amp Rating: {distributionBreaker.AmpRating}A");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Poles:      {distributionBreaker.NumPoles}P");
          break;
        case NodeType.Panel:
          ElectricalEntity.Panel panel = (ElectricalEntity.Panel)entity;
          InfoGroupBox.Text = "Panel " + panel.Name;
          InfoTextBox.AppendText($"Location: {GetLocationString(panel)}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Fed From: {GetParentName(panel.ParentId)}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Distance: {panel.ParentDistance}'");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText("---------------------Panel----------------------");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Bus:     {panel.BusAmpRating}A");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText(
            $"Main:    " + (panel.IsMlo ? "M.L.O." : panel.MainAmpRating + "A")
          );
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Mount:   " + (panel.IsRecessed ? "Recessed" : "Surface"));
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Voltage: {panel.Voltage.Replace(" ", "V-") + "\u03A6"}");
          break;
        case NodeType.PanelBreaker:
          ElectricalEntity.PanelBreaker panelBreaker = (ElectricalEntity.PanelBreaker)entity;
          InfoTextBox.AppendText("--------------------Breaker---------------------");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Amp Rating: {panelBreaker.AmpRating}A");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Poles:      {panelBreaker.NumPoles}P");
          break;
        case NodeType.Disconnect:
          ElectricalEntity.Disconnect disconnect = (ElectricalEntity.Disconnect)entity;
          InfoTextBox.AppendText($"Location: {GetLocationString(disconnect)}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Fed From: {GetParentName(disconnect.ParentId)}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Distance: {disconnect.ParentDistance}'");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText("------------------Disconnect--------------------");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"AS:     {disconnect.AsSize}AS");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"AF:     {disconnect.AfSize}AF");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Poles:  {disconnect.NumPoles}P");
          break;
        case NodeType.Transformer:
          ElectricalEntity.Transformer transformer = (ElectricalEntity.Transformer)entity;
          InfoTextBox.AppendText($"Location: {GetLocationString(transformer)}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Fed From: {GetParentName(transformer.ParentId)}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Distance: {transformer.ParentDistance}'");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText("------------------Transformer-------------------");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"KVA:     {transformer.Kva} KVA");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Voltage: {transformer.Voltage}");
          break;
        case NodeType.Equipment:
          ElectricalEntity.Equipment equipment = (ElectricalEntity.Equipment)entity;
          InfoTextBox.AppendText($"Fed From: {GetParentName(equipment.ParentId)}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText("-------------------Equipment--------------------");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"MCA:      {equipment.Mca}A");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"FLA:      {equipment.Fla}A");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"HP:       {equipment.Hp}HP");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Voltage:  {equipment.Voltage}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Category: {equipment.Category}");
          InfoTextBox.AppendText(Environment.NewLine);
          break;
      }
    }

    private string GetLocationString(PlaceableElectricalEntity entity)
    {
      if (entity.Location.X == 0 && entity.Location.Y == 0)
      {
        return "NOT SET";
      }
      return $"({Math.Round(entity.Location.X, 0)},{Math.Round(entity.Location.Y, 0)})";
    }

    private void TreeView_OnNodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
    {
      ElectricalEntity.ElectricalEntity entity = (ElectricalEntity.ElectricalEntity)e.Node.Tag;
      if (entity != null)
      {
        SetInfoBoxText(entity);
      }
    }

    private void GenerateButton_Click(object sender, EventArgs e)
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
        MakeGroups(startingPoint);
      }
    }

    private GroupType InferGroupType(List<string> groupMembers)
    {
      bool hasMeter = false;
      bool hasMainBreaker = false;
      bool hasDistributionBreaker = false;
      foreach (string groupMember in groupMembers)
      {
        if (serviceList.Select(entity => entity.Id).ToArray().Contains(groupMember))
        {
          return GroupType.PullSection;
        }
        if (meterList.Select(entity => entity.Id).ToArray().Contains(groupMember))
        {
          hasMeter = true;
        }
        if (mainBreakerList.Select(entity => entity.Id).ToArray().Contains(groupMember))
        {
          hasMainBreaker = true;
        }
        if (distributionBreakerList.Select(entity => entity.Id).ToArray().Contains(groupMember))
        {
          hasDistributionBreaker = true;
        }
      }
      if (hasMeter && hasMainBreaker)
      {
        return GroupType.MainMeterAndBreakerSection;
      }
      if (hasDistributionBreaker && hasMeter)
      {
        return GroupType.MultimeterSection;
      }
      if (hasDistributionBreaker)
      {
        return GroupType.DistributionSection;
      }
      if (!hasMeter && hasMainBreaker)
      {
        return GroupType.MainBreakerSection;
      }
      return GroupType.MainMeterSection;
    }

    private void MakeEntityFromDistributionBus(TreeNode distributionBusChild, Point3d currentPoint)
    {
      if (distributionBusChild == null)
      {
        return;
      }
      ElectricalEntity.ElectricalEntity distributionBusChildEntity =
        (ElectricalEntity.ElectricalEntity)distributionBusChild.Tag;
      if (
        distributionBusChild.Nodes.Count == 0
        && distributionBusChildEntity.NodeType == NodeType.Meter
      )
      {
        Meter meter = (Meter)distributionBusChildEntity;
        if (meter.HasCts)
        {
          SingleLine.MakeDistributionCtsMeterCombo(meter, currentPoint);
        }
        else
        {
          SingleLine.MakeDistributionMeterCombo(meter, currentPoint);
        }
      }
      else if (
        distributionBusChild.Nodes.Count == 0
        && distributionBusChildEntity.NodeType == NodeType.DistributionBreaker
      )
      {
        DistributionBreaker distributionBreaker = (DistributionBreaker)distributionBusChildEntity;
        SingleLine.MakeDistributionBreakerCombo(distributionBreaker, currentPoint);
      }
      else if (distributionBusChild.Nodes.Count == 0)
      {
        return;
      }
      else if (distributionBusChildEntity.NodeType == NodeType.Meter)
      {
        Meter meter = (Meter)distributionBusChildEntity;
        TreeNode childNode = distributionBusChild.Nodes[0];
        ElectricalEntity.DistributionBreaker distributionBreaker =
          (ElectricalEntity.DistributionBreaker)childNode.Tag;
        if (meter.HasCts)
        {
          SingleLine.MakeDistributionCtsMeterAndBreakerCombo(
            meter,
            distributionBreaker,
            currentPoint
          );
        }
        else
        {
          SingleLine.MakeDistributionMeterAndBreakerCombo(meter, distributionBreaker, currentPoint);
        }
        if (childNode.Nodes.Count > 0)
        {
          MakeFieldEntity(
            distributionBusChild.Nodes[0],
            new Point3d(currentPoint.X, currentPoint.Y - 4.1875, 0)
          );
          SingleLine.MakeDistributionChildConduit(
            new Point3d(currentPoint.X, currentPoint.Y - 1.6875, 0),
            distributionBreaker.IsExisting()
          );
        }
      }
      else
      {
        DistributionBreaker distributionBreaker = (DistributionBreaker)distributionBusChildEntity;
        SingleLine.MakeDistributionBreakerCombo(distributionBreaker, currentPoint);
        MakeFieldEntity(
          distributionBusChild,
          new Point3d(currentPoint.X, currentPoint.Y - 4.1875, 0)
        );
        SingleLine.MakeDistributionChildConduit(
          new Point3d(currentPoint.X, currentPoint.Y - 1.6875, 0),
          distributionBreaker.IsExisting()
        );
      }
    }

    private List<ElectricalEntity.PanelBreaker> GetPanelBreakersFromPanel(TreeNode panelNode)
    {
      List<ElectricalEntity.PanelBreaker> panelBreakers = new List<PanelBreaker>();
      foreach (TreeNode childNode in panelNode.Nodes)
      {
        ElectricalEntity.ElectricalEntity childEntity = (ElectricalEntity.ElectricalEntity)
          childNode.Tag;
        if (childEntity != null && childEntity.NodeType == NodeType.PanelBreaker)
        {
          ElectricalEntity.PanelBreaker panelBreaker = (ElectricalEntity.PanelBreaker)childEntity;
          panelBreakers.Add(panelBreaker);
        }
      }
      return panelBreakers;
    }

    private void MakeFieldEntity(TreeNode parentNode, Point3d currentPoint)
    {
      if (parentNode == null || parentNode.Nodes.Count == 0)
      {
        return;
      }
      ElectricalEntity.ElectricalEntity parentEntity = (ElectricalEntity.ElectricalEntity)
        parentNode.Tag;
      TreeNode childNode = parentNode.Nodes[0];
      ElectricalEntity.ElectricalEntity childEntity = (ElectricalEntity.ElectricalEntity)
        childNode.Tag;

      if (childEntity.NodeType == NodeType.Panel)
      {
        // Make panel
        ElectricalEntity.Panel panel = (ElectricalEntity.Panel)childEntity;
        SingleLine.MakePanel(panel, currentPoint);

        (string feederWireSize, int feederWireCount) = SingleLine.AddConduitSpec(
          panel,
          currentPoint
        );
        double aicRating;
        if (parentEntity.NodeType == NodeType.Transformer)
        {
          ElectricalEntity.Transformer transformer = (ElectricalEntity.Transformer)parentNode.Tag;
          aicRating = CADObjectCommands.GetAicRatingFromTransformer(
            transformer.Kva,
            1,
            0.035,
            panel.ParentDistance + 10,
            feederWireCount,
            panel.LineVoltage,
            feederWireSize,
            panel.Phase == 3
          );
        }
        else
        {
          aicRating = CADObjectCommands.GetAicRating(
            parentEntity.AicRating,
            panel.ParentDistance + 10,
            feederWireCount,
            panel.LineVoltage,
            feederWireSize,
            panel.Phase == 3
          );
        }

        panel.AicRating = Math.Round(aicRating, 0);
        SingleLine.MakeAicRating(aicRating, currentPoint);
        List<ElectricalEntity.PanelBreaker> panelBreakers = GetPanelBreakersFromPanel(childNode);
        Point3d panelPoint = currentPoint;
        for (int i = 0; i < panelBreakers.Count; i++)
        {
          currentPoint = panelPoint;
          if (i == 0)
          {
            Point3d breakerPoint = new Point3d(currentPoint.X + 0.3125, currentPoint.Y - 0.9333, 0);
            SingleLine.MakeRightPanelBreaker(panelBreakers[i], breakerPoint);
            currentPoint = SingleLine.MakePanelChildConduit(
              i,
              breakerPoint,
              panelBreakers[i].IsExisting()
            );
            TreeNode breakerNode = SingleLineTreeView.Nodes.Find(panelBreakers[i].Id, true)[0];
            MakeFieldEntity(breakerNode, currentPoint);
          }
          if (i == 1)
          {
            Point3d breakerPoint = new Point3d(currentPoint.X - 0.3125, currentPoint.Y - 0.9333, 0);
            SingleLine.MakeLeftPanelBreaker(panelBreakers[i], breakerPoint);
            currentPoint = SingleLine.MakePanelChildConduit(
              i,
              breakerPoint,
              panelBreakers[i].IsExisting()
            );
            TreeNode breakerNode = SingleLineTreeView.Nodes.Find(panelBreakers[i].Id, true)[0];
            MakeFieldEntity(breakerNode, currentPoint);
          }
          if (i == 2)
          {
            Point3d breakerPoint = new Point3d(currentPoint.X + 0.3125, currentPoint.Y - 0.2333, 0);
            SingleLine.MakeRightPanelBreaker(panelBreakers[i], breakerPoint);
            currentPoint = SingleLine.MakePanelChildConduit(
              i,
              breakerPoint,
              panelBreakers[i].IsExisting()
            );
            TreeNode breakerNode = SingleLineTreeView.Nodes.Find(panelBreakers[i].Id, true)[0];
            MakeFieldEntity(breakerNode, currentPoint);
          }
          if (i == 3)
          {
            Point3d breakerPoint = new Point3d(currentPoint.X - 0.3125, currentPoint.Y - 0.2333, 0);
            SingleLine.MakeLeftPanelBreaker(panelBreakers[i], breakerPoint);
            currentPoint = SingleLine.MakePanelChildConduit(
              i,
              breakerPoint,
              panelBreakers[i].IsExisting()
            );
            TreeNode breakerNode = SingleLineTreeView.Nodes.Find(panelBreakers[i].Id, true)[0];
            MakeFieldEntity(breakerNode, currentPoint);
          }
        }
      }
      if (childEntity.NodeType == NodeType.Disconnect)
      {
        ElectricalEntity.Disconnect disconnect = (ElectricalEntity.Disconnect)childEntity;
        SingleLine.MakeDisconnect(disconnect, currentPoint);
        (string feederWireSize, int feederWireCount) = SingleLine.AddConduitSpec(
          disconnect,
          currentPoint
        );
        double aicRating;
        if (parentEntity.NodeType == NodeType.Transformer)
        {
          ElectricalEntity.Transformer transformer = (ElectricalEntity.Transformer)parentNode.Tag;
          aicRating = CADObjectCommands.GetAicRatingFromTransformer(
            transformer.Kva,
            1,
            0.035,
            disconnect.ParentDistance + 10,
            feederWireCount,
            disconnect.LineVoltage,
            feederWireSize,
            disconnect.Phase == 3
          );
        }
        else
        {
          aicRating = CADObjectCommands.GetAicRating(
            parentEntity.AicRating,
            disconnect.ParentDistance + 10,
            feederWireCount,
            disconnect.LineVoltage,
            feederWireSize,
            disconnect.Phase == 3
          );
        }
        disconnect.AicRating = Math.Round(aicRating, 0);

        SingleLine.MakeAicRating(aicRating, currentPoint);

        currentPoint = new Point3d(currentPoint.X, currentPoint.Y - 0.1201, 0);
        if (childNode.Nodes.Count > 0)
        {
          currentPoint = SingleLine.MakeConduitFromDisconnect(
            currentPoint,
            disconnect.IsExisting()
          );
          MakeFieldEntity(childNode, currentPoint);
        }
      }
      if (childEntity.NodeType == NodeType.Transformer)
      {
        ElectricalEntity.Transformer transformer = (ElectricalEntity.Transformer)childEntity;
        SingleLine.MakeTransformer(transformer, currentPoint);
        (string feederWireSize, int feederWireCount) = SingleLine.AddConduitSpec(
          transformer,
          currentPoint
        );

        double aicRating = CADObjectCommands.GetAicRating(
          parentEntity.AicRating,
          transformer.ParentDistance,
          feederWireCount,
          transformer.LineVoltage,
          feederWireSize,
          transformer.Phase == 3
        );
        transformer.AicRating = Math.Round(aicRating, 0);

        SingleLine.MakeAicRating(aicRating, currentPoint);

        currentPoint = new Point3d(currentPoint.X, currentPoint.Y - 0.3739, 0);
        if (childNode.Nodes.Count > 0)
        {
          currentPoint = SingleLine.MakeConduitFromTransformer(
            currentPoint,
            transformer.IsExisting()
          );
          MakeFieldEntity(childNode, currentPoint);
        }
      }
      if (childEntity.NodeType == NodeType.Equipment)
      {
        ElectricalEntity.Equipment equipment = (ElectricalEntity.Equipment)childEntity;
        SingleLine.MakeEquipment(equipment, currentPoint);

        (string feederWireSize, int feederWireCount) = SingleLine.AddConduitSpec(
          equipment,
          currentPoint
        );
        double aicRating = CADObjectCommands.GetAicRating(
          parentEntity.AicRating,
          equipment.ParentDistance,
          feederWireCount,
          equipment.LineVoltage,
          feederWireSize,
          equipment.Is3Phase
        );
      }
    }

    private string MakeDistributionSection(
      string groupId,
      Point3d groupPoint,
      bool isMultimeter,
      string sectionName,
      string distributionBusId = ""
    )
    {
      Point3d busBarPoint = new Point3d(groupPoint.X + 0.25, groupPoint.Y - 0.25, 0);
      Point3d distributionBusChildPoint = new Point3d(
        groupPoint.X,
        groupPoint.Y - 0.0625 - 0.25,
        0
      );
      bool addBusBar = false;
      if (String.IsNullOrEmpty(distributionBusId))
      {
        foreach (string groupMember in groupDict[groupId])
        {
          if (distributionBusList.Select(entity => entity.Id).ToArray().Contains(groupMember))
          {
            distributionBusId = groupMember;
            addBusBar = true;
          }
        }
      }
      double totalBusBarWidth = 0;
      TreeNode distributionBusNode = SingleLineTreeView.Nodes.Find(distributionBusId, true)[0];
      bool groupMembersAdded = false;
      foreach (TreeNode distributionBusChild in distributionBusNode.Nodes)
      {
        double width = 0;
        NodeType nodeType = NodeType.DistributionBreaker;
        if (!groupDict[groupId].Contains(distributionBusChild.Name))
        {
          width = 2 + AggregateEntityWidth(distributionBusChild, nodeType);
          totalBusBarWidth += width;
          continue;
        }
        if (meterList.Select(entity => entity.Id).ToArray().Contains(distributionBusChild.Name))
        {
          nodeType = NodeType.Meter;
        }
        width = 2 + AggregateEntityWidth(distributionBusChild, nodeType);
        totalBusBarWidth += width;
        // move to center of width
        distributionBusChildPoint = new Point3d(
          distributionBusChildPoint.X + (width / 2),
          distributionBusChildPoint.Y,
          distributionBusChildPoint.Z
        );
        MakeEntityFromDistributionBus(distributionBusChild, distributionBusChildPoint);
        // move to end of width
        distributionBusChildPoint = new Point3d(
          distributionBusChildPoint.X + (width / 2),
          distributionBusChildPoint.Y,
          distributionBusChildPoint.Z
        );
        groupMembersAdded = true;
      }
      if (!groupMembersAdded)
      {
        // this means we've encountered another bus bar
        return MakeDistributionSection(groupId, groupPoint, isMultimeter, sectionName);
      }
      if (addBusBar)
      {
        ElectricalEntity.DistributionBus distributionBus = (ElectricalEntity.DistributionBus)
          distributionBusNode.Tag;
        SingleLine.MakeDistributionBus(
          distributionBus,
          isMultimeter,
          totalBusBarWidth,
          groupPoint,
          sectionName
        );
      }
      else
      {
        // todo: add label above section
      }
      return distributionBusId;
    }

    private ElectricalEntity.Meter GetMeterFromMainSection(string groupId)
    {
      foreach (string entityId in groupDict[groupId])
      {
        foreach (ElectricalEntity.Meter meter in meterList)
        {
          if (meter.Id == entityId)
          {
            return meter;
          }
        }
      }
      ElectricalEntity.Meter nullMeter = new ElectricalEntity.Meter(
        "",
        "",
        "",
        false,
        false,
        0,
        new Point()
      );
      nullMeter.NodeType = NodeType.Undefined;
      return nullMeter;
    }

    private ElectricalEntity.MainBreaker GetMainBreakerFromMainSection(string groupId)
    {
      foreach (string entityId in groupDict[groupId])
      {
        foreach (ElectricalEntity.MainBreaker mainBreaker in mainBreakerList)
        {
          if (mainBreaker.Id == entityId)
          {
            return mainBreaker;
          }
        }
      }
      ElectricalEntity.MainBreaker nullMainBreaker = new ElectricalEntity.MainBreaker(
        "",
        "",
        "",
        0,
        false,
        false,
        0,
        0,
        new Point()
      );
      nullMainBreaker.NodeType = NodeType.Undefined;
      return nullMainBreaker;
    }

    private ElectricalEntity.DistributionBus GetDistributionBusFromDistributionSection(
      string groupId
    )
    {
      foreach (string entityId in groupDict[groupId])
      {
        foreach (ElectricalEntity.DistributionBus distributionBus in distributionBusList)
        {
          if (distributionBus.Id == entityId)
          {
            return distributionBus;
          }
        }
      }
      ElectricalEntity.DistributionBus nullDistributionBus = new ElectricalEntity.DistributionBus(
        "",
        "",
        "",
        0,
        0,
        new Point(),
        new Point3d()
      );
      nullDistributionBus.NodeType = NodeType.Undefined;
      return nullDistributionBus;
    }

    private string GetGroupNameFromId(string groupId)
    {
      string name = string.Empty;
      GroupNode group = groupList.Find(groupList => groupList.Id == groupId);
      if (group == null)
      {
        name = group.Name;
      }
      return name;
    }

    private ElectricalEntity.Service GetServiceFromPullSection(string groupId)
    {
      foreach (string entityId in groupDict[groupId])
      {
        foreach (ElectricalEntity.Service service in serviceList)
        {
          if (service.Id == entityId)
          {
            return service;
          }
        }
      }
      ElectricalEntity.Service nullService = new ElectricalEntity.Service(
        "",
        "",
        "",
        0,
        "",
        0,
        new Point(),
        new Point3d()
      );
      nullService.NodeType = NodeType.Undefined;
      return nullService;
    }

    private void MakeGroups(Point3d startingPoint)
    {
      Point3d currentPoint = startingPoint;
      int index = 1;
      string distributionBusId = String.Empty;
      ElectricalEntity.DistributionBus distributionBus;
      bool existing = false;
      foreach (string groupId in groupDict.Keys)
      {
        if (String.IsNullOrEmpty(groupId))
          continue;
        GroupType groupType = InferGroupType(groupDict[groupId]);
        double groupWidth = 2;
        if (groupType == GroupType.MultimeterSection || groupType == GroupType.DistributionSection)
        {
          groupWidth = AggregateGroupWidth(groupId);
          string groupName = GetGroupNameFromId(groupId);
          distributionBusId = MakeDistributionSection(
            groupId,
            currentPoint,
            groupType == GroupType.MultimeterSection,
            groupName,
            distributionBusId
          );
          ElectricalEntity.DistributionBus thisDistributionBus =
            GetDistributionBusFromDistributionSection(groupId);
          if (thisDistributionBus.NodeType != NodeType.Undefined)
          {
            distributionBus = thisDistributionBus;
            existing = distributionBus.Status.ToLower() == "existing";
          }
        }
        else if (groupType == GroupType.PullSection)
        {
          groupWidth = 0.75;
          ElectricalEntity.Service service = GetServiceFromPullSection(groupId);
          existing = service.Status.ToLower() == "existing";
          SingleLine.MakePullSection(service, currentPoint);
        }
        else if (groupType == GroupType.MainMeterSection)
        {
          // get meter from section
          ElectricalEntity.Meter meter = GetMeterFromMainSection(groupId);
          existing = meter.Status.ToLower() == "existing";
          groupWidth = 1.75;
          SingleLine.MakeMainMeterSection(meter, currentPoint);
        }
        else if (groupType == GroupType.MainBreakerSection)
        {
          // get breaker from section
          ElectricalEntity.MainBreaker mainBreaker = GetMainBreakerFromMainSection(groupId);
          existing = mainBreaker.Status.ToLower() == "existing";
          groupWidth = 1.75;
          SingleLine.MakeMainBreakerSection(mainBreaker, currentPoint);
        }
        else if (groupType == GroupType.MainMeterAndBreakerSection)
        {
          // get meter and breaker from section
          ElectricalEntity.Meter meter = GetMeterFromMainSection(groupId);
          ElectricalEntity.MainBreaker mainBreaker = GetMainBreakerFromMainSection(groupId);
          existing =
            (meter.Status.ToLower() == "existing") && (mainBreaker.Status.ToLower() == "existing");
          groupWidth = 1.75;
          SingleLine.MakeMainMeterAndBreakerSection(meter, mainBreaker, currentPoint);
        }

        // draw lines on CAD
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
          LineData boxLine1 = new LineData();
          if (existing)
          {
            boxLine1.ColorIndex = 8;
          }
          boxLine1.Layer = "E-SYM1";
          boxLine1.StartPoint = new SimpleVector3d();
          boxLine1.EndPoint = new SimpleVector3d();
          boxLine1.StartPoint.X = currentPoint.X;
          boxLine1.StartPoint.Y = currentPoint.Y;
          boxLine1.EndPoint.X = currentPoint.X + groupWidth;
          boxLine1.EndPoint.Y = currentPoint.Y;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, boxLine1, 1, "HIDDEN");

          LineData boxLine2 = new LineData();
          boxLine2.Layer = "E-SYM1";
          if (existing)
          {
            boxLine2.ColorIndex = 8;
          }
          boxLine2.StartPoint = new SimpleVector3d();
          boxLine2.EndPoint = new SimpleVector3d();
          boxLine2.StartPoint.X = currentPoint.X;
          boxLine2.StartPoint.Y = currentPoint.Y;
          boxLine2.EndPoint.X = currentPoint.X;
          boxLine2.EndPoint.Y = currentPoint.Y - 2;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, boxLine2, 1, "HIDDEN");

          LineData boxLine3 = new LineData();
          boxLine3.Layer = "E-SYM1";
          if (existing)
          {
            boxLine3.ColorIndex = 8;
          }
          boxLine3.StartPoint = new SimpleVector3d();
          boxLine3.EndPoint = new SimpleVector3d();
          boxLine3.StartPoint.X = currentPoint.X;
          boxLine3.StartPoint.Y = currentPoint.Y - 2;
          boxLine3.EndPoint.X = currentPoint.X + groupWidth;
          boxLine3.EndPoint.Y = currentPoint.Y - 2;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, boxLine3, 1, "HIDDEN");

          if (index == groupDict.Keys.Count - 1)
          {
            // draw last vertical line
            LineData boxLine4 = new LineData();
            boxLine4.Layer = "E-SYM1";
            if (existing)
            {
              boxLine4.ColorIndex = 8;
            }
            boxLine4.StartPoint = new SimpleVector3d();
            boxLine4.EndPoint = new SimpleVector3d();
            boxLine4.StartPoint.X = currentPoint.X + groupWidth;
            boxLine4.StartPoint.Y = currentPoint.Y;
            boxLine4.EndPoint.X = currentPoint.X + groupWidth;
            boxLine4.EndPoint.Y = currentPoint.Y - 2;
            CADObjectCommands.CreateLine(new Point3d(), tr, btr, boxLine4, 1, "HIDDEN");
          }
          else
          {
            index++;
          }
          tr.Commit();
        }
        currentPoint = new Point3d(currentPoint.X + groupWidth, currentPoint.Y, 0);
      }
    }

    private void SingleLineDialogWindow_Load(object sender, EventArgs e) { }

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
      foreach (PlaceableElectricalEntity placeable in placeables)
      {
        foreach (PlaceableElectricalEntity pooledPlaceable in pooledEquipment)
        {
          if (placeable.Id == pooledPlaceable.Id)
          {
            placeable.ParentDistance = pooledPlaceable.ParentDistance;
          }
        }
      }
    }

    private void PlaceOnPlanButton_Click(object sender, EventArgs e)
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
        foreach (PlaceableElectricalEntity placeable in placeables)
        {
          Point3d? p = placeable.Place();
          if (p == null)
          {
            placeable.Location = new Point3d(0, 0, 0);
            break;
          }
        }
      }
      CalculateDistances();

      SingleLineTreeView.BeginUpdate();
      SingleLineTreeView.Nodes.Clear();
      PopulateTreeView();
      SingleLineTreeView.ExpandAll();
      SingleLineTreeView.EndUpdate();
      if (serviceList.Count > 0)
      {
        SetInfoBoxText(serviceList[0]);
      }
      placeables.ForEach(gmepDb.UpdatePlaceable);
      foreach (TreeNode node in SingleLineTreeView.Nodes)
      {
        if (node != null)
        {
          ContextMenuStrip docMenu = new ContextMenuStrip();
          ToolStripMenuItem placeLabel = new ToolStripMenuItem();
          placeLabel.Text = "Place";
          placeLabel.Click += TreeViewPlaceSelected_Click;
          docMenu.Items.Add(placeLabel);
          node.ContextMenuStrip = docMenu;
          SetTreeContextMenu(node);
        }
      }
    }

    public void TreeViewPlaceSelected_Click(object sender, EventArgs e)
    {
      if (SingleLineTreeView.SelectedNode == null)
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
        ElectricalEntity.PlaceableElectricalEntity placeable =
          (ElectricalEntity.PlaceableElectricalEntity)SingleLineTreeView.SelectedNode.Tag;
        Point3d? p = placeable.Place();
        if (p == null)
        {
          placeable.Location = new Point3d(0, 0, 0);
        }
      }
      CalculateDistances();

      SingleLineTreeView.BeginUpdate();
      SingleLineTreeView.Nodes.Clear();
      PopulateTreeView();
      SingleLineTreeView.ExpandAll();
      SingleLineTreeView.EndUpdate();
      if (serviceList.Count > 0)
      {
        SetInfoBoxText(serviceList[0]);
      }
      placeables.ForEach(gmepDb.UpdatePlaceable);
      foreach (TreeNode node in SingleLineTreeView.Nodes)
      {
        if (node != null)
        {
          ContextMenuStrip docMenu = new ContextMenuStrip();
          ToolStripMenuItem placeLabel = new ToolStripMenuItem();
          placeLabel.Text = "Place";
          placeLabel.Click += TreeViewPlaceSelected_Click;
          docMenu.Items.Add(placeLabel);
          node.ContextMenuStrip = docMenu;
          SetTreeContextMenu(node);
        }
      }
    }
  }
}
