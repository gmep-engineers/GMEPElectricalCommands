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

    public void PopulateTreeView()
    {
      foreach (ElectricalEntity.Service service in serviceList)
      {
        TreeNode serviceNode = SingleLineTreeView.Nodes.Add(service.Id, service.Name);
        serviceNode.Tag = service;
        SetTreeNodeColor(serviceNode, service);
        PopulateFromService(serviceNode, service.NodeId);
      }
    }

    public void PopulateFromService(TreeNode node, string serviceNodeId)
    {
      foreach (ElectricalEntity.Meter meter in meterList)
      {
        if (VerifyNodeLink(serviceNodeId, meter.NodeId))
        {
          TreeNode meterNode = node.Nodes.Add(meter.Id, meter.Name);
          meterNode.Tag = meter;
          PopulateFromMainMeter(meterNode, meter.NodeId);
        }
      }
      foreach (ElectricalEntity.MainBreaker mainBreaker in mainBreakerList)
      {
        if (VerifyNodeLink(serviceNodeId, mainBreaker.NodeId))
        {
          TreeNode mainBreakerNode = node.Nodes.Add(mainBreaker.Id, mainBreaker.Name);
          mainBreakerNode.Tag = mainBreaker;
          PopulateFromMainBreaker(mainBreakerNode, mainBreaker.NodeId);
        }
      }
    }

    public void PopulateFromMainMeter(TreeNode node, string meterNodeId)
    {
      foreach (ElectricalEntity.MainBreaker mainBreaker in mainBreakerList)
      {
        if (VerifyNodeLink(meterNodeId, mainBreaker.NodeId))
        {
          TreeNode mainBreakerNode = node.Nodes.Add(mainBreaker.Id, mainBreaker.Name);
          mainBreakerNode.Tag = mainBreaker;
          PopulateFromMainBreaker(mainBreakerNode, mainBreaker.NodeId);
        }
      }
    }

    public void PopulateFromMainBreaker(TreeNode node, string mainBreakerNodeId)
    {
      foreach (ElectricalEntity.DistributionBus distributionBus in distributionBusList)
      {
        if (VerifyNodeLink(mainBreakerNodeId, distributionBus.NodeId))
        {
          TreeNode distributionBusNode = node.Nodes.Add(distributionBus.Id, distributionBus.Name);
          distributionBusNode.Tag = distributionBus;
          SetTreeNodeColor(distributionBusNode, distributionBus);
          PopulateFromDistributionBus(distributionBusNode, distributionBus.NodeId);
        }
      }
    }

    public void PopulateFromDistributionBus(TreeNode node, string distributionBusNodeId)
    {
      foreach (ElectricalEntity.Meter meter in meterList)
      {
        if (VerifyNodeLink(distributionBusNodeId, meter.NodeId))
        {
          TreeNode distributionMeterNode = node.Nodes.Add(meter.Id, meter.Name);
          distributionMeterNode.Tag = meter;
          PopulateFromDistributionMeter(distributionMeterNode, meter.NodeId);
        }
      }
      foreach (ElectricalEntity.DistributionBreaker distributionBreaker in distributionBreakerList)
      {
        if (VerifyNodeLink(distributionBusNodeId, distributionBreaker.NodeId))
        {
          TreeNode distributionBreakerNode = node.Nodes.Add(
            distributionBreaker.Id,
            distributionBreaker.Name
          );
          distributionBreakerNode.Tag = distributionBreaker;
          PopulateFromDistributionBreaker(distributionBreakerNode, distributionBreaker.NodeId);
        }
      }
    }

    public void PopulateFromDistributionMeter(TreeNode node, string meterNodeId)
    {
      foreach (ElectricalEntity.DistributionBreaker distributionBreaker in distributionBreakerList)
      {
        if (VerifyNodeLink(meterNodeId, distributionBreaker.NodeId))
        {
          TreeNode distributionBreakerNode = node.Nodes.Add(
            distributionBreaker.Id,
            distributionBreaker.Name
          );
          distributionBreakerNode.Tag = distributionBreaker;
          PopulateFromDistributionBreaker(distributionBreakerNode, distributionBreaker.NodeId);
        }
      }
    }

    public void PopulateFromDistributionBreaker(TreeNode node, string distributionBreakerNodeId)
    {
      foreach (ElectricalEntity.Panel panel in panelList)
      {
        if (VerifyNodeLink(distributionBreakerNodeId, panel.NodeId))
        {
          TreeNode panelNode = node.Nodes.Add(panel.Id, panel.Name);
          panelNode.Tag = panel;
          SetTreeNodeColor(panelNode, panel);
          PopulateFromPanel(panelNode, panel.NodeId);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(distributionBreakerNodeId, disconnect.NodeId))
        {
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          disconnectNode.Tag = disconnect;
          SetTreeNodeColor(disconnectNode, disconnect);
          PopulateFromDisconnect(disconnectNode, disconnect.NodeId);
        }
      }
      foreach (ElectricalEntity.Transformer transformer in transformerList)
      {
        if (VerifyNodeLink(distributionBreakerNodeId, transformer.NodeId))
        {
          TreeNode transformerNode = node.Nodes.Add(transformer.Id, transformer.Name);
          transformerNode.Tag = transformer;
          SetTreeNodeColor(transformerNode, transformer);
          PopulateFromTransformer(transformerNode, transformer.NodeId);
        }
      }
    }

    public void PopulateFromPanel(TreeNode node, string panelNodeId)
    {
      foreach (ElectricalEntity.PanelBreaker panelBreaker in panelBreakerList)
      {
        if (VerifyNodeLink(panelNodeId, panelBreaker.NodeId))
        {
          TreeNode panelBreakerNode = node.Nodes.Add(panelBreaker.Id, panelBreaker.Name);
          panelBreakerNode.Tag = panelBreaker;
          PopulateFromPanelBreaker(panelBreakerNode, panelBreaker.NodeId);
        }
      }
      foreach (ElectricalEntity.Panel panel in panelList)
      {
        if (VerifyNodeLink(panelNodeId, panel.NodeId))
        {
          TreeNode panelNode = node.Nodes.Add(panel.Id, panel.Name);
          panelNode.Tag = panel;
          SetTreeNodeColor(panelNode, panel);
          PopulateFromPanel(panelNode, panel.NodeId);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(panelNodeId, disconnect.NodeId))
        {
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          disconnectNode.Tag = disconnect;
          SetTreeNodeColor(disconnectNode, disconnect);
          PopulateFromDisconnect(disconnectNode, disconnect.NodeId);
        }
      }
      foreach (ElectricalEntity.Transformer transformer in transformerList)
      {
        if (VerifyNodeLink(panelNodeId, transformer.NodeId))
        {
          TreeNode transformerNode = node.Nodes.Add(transformer.Id, transformer.Name);
          transformerNode.Tag = transformer;
          SetTreeNodeColor(transformerNode, transformer);
          PopulateFromTransformer(transformerNode, transformer.NodeId);
        }
      }
    }

    public void PopulateFromPanelBreaker(TreeNode node, string panelBreakerNodeId)
    {
      foreach (ElectricalEntity.Panel panel in panelList)
      {
        if (VerifyNodeLink(panelBreakerNodeId, panel.NodeId))
        {
          TreeNode panelNode = node.Nodes.Add(panel.Id, panel.Name);
          panelNode.Tag = panel;
          SetTreeNodeColor(panelNode, panel);
          PopulateFromPanel(panelNode, panel.NodeId);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(panelBreakerNodeId, disconnect.NodeId))
        {
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          disconnectNode.Tag = disconnect;
          SetTreeNodeColor(disconnectNode, disconnect);
          PopulateFromDisconnect(disconnectNode, disconnect.NodeId);
        }
      }
      foreach (ElectricalEntity.Transformer transformer in transformerList)
      {
        if (VerifyNodeLink(panelBreakerNodeId, transformer.NodeId))
        {
          TreeNode transformerNode = node.Nodes.Add(transformer.Id, transformer.Name);
          transformerNode.Tag = transformer;
          SetTreeNodeColor(transformerNode, transformer);
          PopulateFromTransformer(transformerNode, transformer.NodeId);
        }
      }
    }

    public void PopulateFromDisconnect(TreeNode node, string disconnectNodeId)
    {
      foreach (ElectricalEntity.Panel panel in panelList)
      {
        if (VerifyNodeLink(disconnectNodeId, panel.NodeId))
        {
          TreeNode panelNode = node.Nodes.Add(panel.Id, panel.Name);
          panelNode.Tag = panel;
          SetTreeNodeColor(panelNode, panel);
          PopulateFromPanel(panelNode, panel.NodeId);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(disconnectNodeId, disconnect.NodeId))
        {
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          disconnectNode.Tag = disconnect;
          SetTreeNodeColor(disconnectNode, disconnect);
          PopulateFromDisconnect(disconnectNode, disconnect.NodeId);
        }
      }
      foreach (ElectricalEntity.Transformer transformer in transformerList)
      {
        if (VerifyNodeLink(disconnectNodeId, transformer.NodeId))
        {
          TreeNode transformerNode = node.Nodes.Add(transformer.Id, transformer.Name);
          transformerNode.Tag = transformer;
          SetTreeNodeColor(transformerNode, transformer);
          PopulateFromTransformer(transformerNode, transformer.NodeId);
        }
      }
    }

    public void PopulateFromTransformer(TreeNode node, string transformerNodeId)
    {
      foreach (ElectricalEntity.Panel panel in panelList)
      {
        if (VerifyNodeLink(transformerNodeId, panel.NodeId))
        {
          TreeNode panelNode = node.Nodes.Add(panel.Id, panel.Name);
          panelNode.Tag = panel;
          SetTreeNodeColor(panelNode, panel);
          PopulateFromPanel(panelNode, panel.NodeId);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(transformerNodeId, disconnect.NodeId))
        {
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          disconnectNode.Tag = disconnect;
          SetTreeNodeColor(disconnectNode, disconnect);
          PopulateFromDisconnect(disconnectNode, disconnect.NodeId);
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

    private void SetInfoBoxText(ElectricalEntity.ElectricalEntity entity)
    {
      InfoTextBox.Clear();
      InfoGroupBox.Text = entity.Name;
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
          InfoTextBox.AppendText($"Voltage:    {service.Voltage.Replace(" ", "V-")}");
          break;
        case NodeType.Meter:
          ElectricalEntity.Meter meter = (ElectricalEntity.Meter)entity;
          InfoTextBox.AppendText("---------------------Meter----------------------");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"ID:     {meter.Id}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Status: {meter.Status}");
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
          InfoTextBox.AppendText($"Location: {GetLocationString(panel)}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText("---------------------Panel----------------------");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Bus:     {panel.BusAmpRating}A");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText(
            $"Main:    " + (panel.IsMlo ? "M.L.O." : panel.MainAmpRating + "A")
          );
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Voltage: {panel.Voltage.Replace(" ", "V-")}");
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
          InfoTextBox.AppendText("------------------Transformer-------------------");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"KVA:     {transformer.Kva} KVA");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Voltage: {transformer.Voltage}" + "\u0081");
          break;
      }
    }

    private string GetLocationString(PlaceableElectricalEntity entity)
    {
      if (entity.Location.X == 0 && entity.Location.Y == 0)
      {
        return "NOT SET";
      }
      return $"{Math.Round(entity.Location.X, 0)},{Math.Round(entity.Location.Y, 0)}";
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
            new Point3d(currentPoint.X, currentPoint.Y - 4, 0)
          );
          SingleLine.MakeDistributionChildConduit(
            new Point3d(currentPoint.X, currentPoint.Y - 1.6875, 0)
          );
        }
      }
      else
      {
        DistributionBreaker distributionBreaker = (DistributionBreaker)distributionBusChildEntity;
        SingleLine.MakeDistributionBreakerCombo(distributionBreaker, currentPoint);
        MakeFieldEntity(distributionBusChild, new Point3d(currentPoint.X, currentPoint.Y - 4, 0));
        SingleLine.MakeDistributionChildConduit(
          new Point3d(currentPoint.X, currentPoint.Y - 1.6875, 0)
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
      TreeNode childNode = parentNode.Nodes[0];
      ElectricalEntity.ElectricalEntity childEntity = (ElectricalEntity.ElectricalEntity)
        childNode.Tag;
      if (childEntity.NodeType == NodeType.Panel)
      {
        // Make panel
        ElectricalEntity.Panel panel = (ElectricalEntity.Panel)childEntity;
        SingleLine.MakePanel(panel, currentPoint);
        List<ElectricalEntity.PanelBreaker> panelBreakers = GetPanelBreakersFromPanel(childNode);
        for (int i = 0; i < panelBreakers.Count; i++)
        {
          if (i == 0)
          {
            Point3d breakerPoint = new Point3d(currentPoint.X + 0.3125, currentPoint.Y - 0.9333, 0);
            SingleLine.MakeRightPanelBreaker(panelBreakers[i], breakerPoint);
            currentPoint = SingleLine.MakePanelChildConduit(i, breakerPoint);
            MakeFieldEntity(childNode.Nodes[i], currentPoint);
          }
          if (i == 1)
          {
            Point3d breakerPoint = new Point3d(currentPoint.X - 0.3125, currentPoint.Y - 0.9333, 0);
            SingleLine.MakeRightPanelBreaker(panelBreakers[i], breakerPoint);
            currentPoint = SingleLine.MakePanelChildConduit(i, breakerPoint);
            MakeFieldEntity(childNode.Nodes[i], currentPoint);
          }
          if (i == 2)
          {
            Point3d breakerPoint = new Point3d(currentPoint.X + 0.3125, currentPoint.Y - 0.1833, 0);
            SingleLine.MakeRightPanelBreaker(panelBreakers[i], breakerPoint);
            currentPoint = SingleLine.MakePanelChildConduit(i, breakerPoint);
            MakeFieldEntity(childNode.Nodes[i], currentPoint);
          }
          if (i == 3)
          {
            Point3d breakerPoint = new Point3d(currentPoint.X - 0.3125, currentPoint.Y - 0.1833, 0);
            SingleLine.MakeRightPanelBreaker(panelBreakers[i], breakerPoint);
            currentPoint = SingleLine.MakePanelChildConduit(i, breakerPoint);
            MakeFieldEntity(childNode.Nodes[i], currentPoint);
          }
        }
      }
      if (childEntity.NodeType == NodeType.Disconnect)
      {
        ElectricalEntity.Disconnect disconnect = (ElectricalEntity.Disconnect)childEntity;
        SingleLine.MakeDisconnect(disconnect, currentPoint);
        currentPoint = new Point3d(currentPoint.X, currentPoint.Y - 0.1201, 0);
        if (childNode.Nodes.Count > 0)
        {
          currentPoint = SingleLine.MakeConduitFromDisconnect(currentPoint);
          MakeFieldEntity(childNode, currentPoint);
        }
      }
      if (childEntity.NodeType == NodeType.Transformer)
      {
        ElectricalEntity.Transformer transformer = (ElectricalEntity.Transformer)childEntity;
        SingleLine.MakeTransformer(transformer, currentPoint);
        currentPoint = new Point3d(currentPoint.X, currentPoint.Y - 0.3739, 0);
        if (childNode.Nodes.Count > 0)
        {
          currentPoint = SingleLine.MakeConduitFromTransformer(currentPoint);
          MakeFieldEntity(childNode, currentPoint);
        }
      }
    }

    private string MakeDistributionSection(
      string groupId,
      Point3d groupPoint,
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
        return MakeDistributionSection(groupId, groupPoint);
      }
      if (addBusBar)
      {
        Polyline2dData polyData = new Polyline2dData();
        polyData.Layer = "E-CND1";
        polyData.Vertices.Add(new SimpleVector3d(busBarPoint.X, busBarPoint.Y, 0));
        polyData.Vertices.Add(
          new SimpleVector3d(busBarPoint.X + totalBusBarWidth - 0.5, busBarPoint.Y, 0)
        );
        polyData.Vertices.Add(
          new SimpleVector3d(busBarPoint.X + totalBusBarWidth - 0.5, busBarPoint.Y - 0.0625, 0)
        );
        polyData.Vertices.Add(new SimpleVector3d(busBarPoint.X, busBarPoint.Y - 0.0625, 0));
        polyData.Vertices.Add(new SimpleVector3d(busBarPoint.X, busBarPoint.Y, 0));
        polyData.Closed = true;
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
          CADObjectCommands.CreatePolyline2d(new Point3d(), tr, btr, polyData, 1);
          tr.Commit();
        }
      }
      return distributionBusId;
    }

    private void MakeGroups(Point3d startingPoint)
    {
      Point3d currentPoint = startingPoint;
      int index = 1;
      string distributionBusId = String.Empty;
      foreach (string groupId in groupDict.Keys)
      {
        if (String.IsNullOrEmpty(groupId))
          continue;
        GroupType groupType = InferGroupType(groupDict[groupId]);

        double groupWidth = 2;
        if (groupType == GroupType.MultimeterSection || groupType == GroupType.DistributionSection)
        {
          groupWidth = AggregateGroupWidth(groupId);
          distributionBusId = MakeDistributionSection(groupId, currentPoint, distributionBusId);
        }
        if (groupType == GroupType.PullSection)
        {
          groupWidth = 0.75;
          // MakePullSection(startingPoint)
        }
        if (groupType == GroupType.MainMeterSection)
        {
          // MakeMainMeterSection(startingPoint);
          // MakeGroundBus(startingPoint);
        }
        if (groupType == GroupType.MainBreakerSection)
        {
          // MakeMainBreakerSection(startingPoint);
          // MakeGroundBus(startingPoint);
        }
        if (groupType == GroupType.MainMeterAndBreakerSection)
        {
          // MakeMainMeterAndBreakerSection(startingPoint);
          // MakeGroundBus(startingPoint);
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
          boxLine2.StartPoint = new SimpleVector3d();
          boxLine2.EndPoint = new SimpleVector3d();
          boxLine2.StartPoint.X = currentPoint.X;
          boxLine2.StartPoint.Y = currentPoint.Y;
          boxLine2.EndPoint.X = currentPoint.X;
          boxLine2.EndPoint.Y = currentPoint.Y - 2;
          CADObjectCommands.CreateLine(new Point3d(), tr, btr, boxLine2, 1, "HIDDEN");

          LineData boxLine3 = new LineData();
          boxLine3.Layer = "E-SYM1";
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
      for (int i = 0; i < pooledEquipment.Count; i++)
      {
        bool isMatch = false;

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
    }
  }
}
