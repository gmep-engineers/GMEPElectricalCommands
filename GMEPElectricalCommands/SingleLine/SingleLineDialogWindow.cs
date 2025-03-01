using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Accord.Statistics.Distributions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
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

    public double AggregateEntityWidth(string entityId, NodeType entityType)
    {
      double width = 0;
      TreeNode[] nodes = SingleLineTreeView.Nodes.Find(entityId, true);
      if (nodes == null || nodes.Length == 0)
      {
        return 0;
      }
      TreeNode thisNode = nodes[0];
      foreach (TreeNode node in thisNode.Nodes)
      {
        if (node.Nodes.Count == 0)
        {
          continue;
        }
        TreeNode childNode = node.Nodes[0];
        ElectricalEntity.ElectricalEntity childEntity = (ElectricalEntity.ElectricalEntity)
          childNode.Tag;
        switch (childEntity.NodeType)
        {
          case NodeType.Panel:
            width +=
              AggregateEntityWidth(childEntity.Id, childEntity.NodeType)
              + (entityType == NodeType.Panel ? 2 : 0);
            break;
          case NodeType.Disconnect:
            width +=
              AggregateEntityWidth(childEntity.Id, childEntity.NodeType)
              + (entityType == NodeType.Panel ? 2 : 0);
            ;
            break;
          case NodeType.Transformer:
            width +=
              AggregateEntityWidth(childEntity.Id, childEntity.NodeType)
              + (entityType == NodeType.Panel ? 2 : 0);
            ;
            break;
          case NodeType.PanelBreaker:
            width += 2 + AggregateEntityWidth(childEntity.Id, childEntity.NodeType);
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
        width += 2 + AggregateEntityWidth(childId, NodeType.Undefined);
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
        width += AggregateEntityWidth(childEntity.Id, childEntity.NodeType);
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

    public void PopulateTreeView()
    {
      foreach (ElectricalEntity.Service service in serviceList)
      {
        TreeNode serviceNode = SingleLineTreeView.Nodes.Add(service.Id, service.Name);
        serviceNode.Tag = service;
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
          PopulateFromPanel(panelNode, panel.NodeId);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(distributionBreakerNodeId, disconnect.NodeId))
        {
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          disconnectNode.Tag = disconnect;
          PopulateFromDisconnect(disconnectNode, disconnect.NodeId);
        }
      }
      foreach (ElectricalEntity.Transformer transformer in transformerList)
      {
        if (VerifyNodeLink(distributionBreakerNodeId, transformer.NodeId))
        {
          TreeNode transformerNode = node.Nodes.Add(transformer.Id, transformer.Name);
          transformerNode.Tag = transformer;
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
          PopulateFromPanel(panelNode, panel.NodeId);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(panelNodeId, disconnect.NodeId))
        {
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          disconnectNode.Tag = disconnect;
          PopulateFromDisconnect(disconnectNode, disconnect.NodeId);
        }
      }
      foreach (ElectricalEntity.Transformer transformer in transformerList)
      {
        if (VerifyNodeLink(panelNodeId, transformer.NodeId))
        {
          TreeNode transformerNode = node.Nodes.Add(transformer.Id, transformer.Name);
          transformerNode.Tag = transformer;
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
          PopulateFromPanel(panelNode, panel.NodeId);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(panelBreakerNodeId, disconnect.NodeId))
        {
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          disconnectNode.Tag = disconnect;
          PopulateFromDisconnect(disconnectNode, disconnect.NodeId);
        }
      }
      foreach (ElectricalEntity.Transformer transformer in transformerList)
      {
        if (VerifyNodeLink(panelBreakerNodeId, transformer.NodeId))
        {
          TreeNode transformerNode = node.Nodes.Add(transformer.Id, transformer.Name);
          transformerNode.Tag = transformer;
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
          PopulateFromPanel(panelNode, panel.NodeId);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(disconnectNodeId, disconnect.NodeId))
        {
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          disconnectNode.Tag = disconnect;
          PopulateFromDisconnect(disconnectNode, disconnect.NodeId);
        }
      }
      foreach (ElectricalEntity.Transformer transformer in transformerList)
      {
        if (VerifyNodeLink(disconnectNodeId, transformer.NodeId))
        {
          TreeNode transformerNode = node.Nodes.Add(transformer.Id, transformer.Name);
          transformerNode.Tag = transformer;
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
          PopulateFromPanel(panelNode, panel.NodeId);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(transformerNodeId, disconnect.NodeId))
        {
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          disconnectNode.Tag = disconnect;
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
      switch (entity.NodeType)
      {
        case NodeType.Service:
          ElectricalEntity.Service service = (ElectricalEntity.Service)entity;
          InfoTextBox.AppendText($"ID:         {service.Id}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Status:     {service.Status}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"AIC:        {service.AicRating} AIC");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Amp Rating: {service.AmpRating}A");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Voltage:    {service.Voltage.Replace(" ", "V-")}");
          break;
        case NodeType.Meter:
          ElectricalEntity.Meter meter = (ElectricalEntity.Meter)entity;
          InfoTextBox.AppendText($"ID:     {meter.Id}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Status: {meter.Status}");
          break;
        case NodeType.MainBreaker:
          ElectricalEntity.MainBreaker mainBreaker = (ElectricalEntity.MainBreaker)entity;
          InfoTextBox.AppendText($"ID:         {mainBreaker.Id}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Status:     {mainBreaker.Status}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"AIC:        {mainBreaker.AicRating} AIC");
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
          InfoTextBox.AppendText($"ID:         {distributionBus.Id}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Status:     {distributionBus.Status}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"AIC:        {distributionBus.AicRating} AIC");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Amp Rating: {distributionBus.AmpRating}A");
          break;
        case NodeType.DistributionBreaker:
          ElectricalEntity.DistributionBreaker distributionBreaker =
            (ElectricalEntity.DistributionBreaker)entity;
          InfoTextBox.AppendText($"ID:         {distributionBreaker.Id}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Status:     {distributionBreaker.Status}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"AIC:        {distributionBreaker.AicRating} AIC");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Amp Rating: {distributionBreaker.AmpRating}A");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Poles:      {distributionBreaker.NumPoles}P");
          break;
        case NodeType.Panel:
          ElectricalEntity.Panel panel = (ElectricalEntity.Panel)entity;
          InfoTextBox.AppendText($"ID:      {panel.Id}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Status:  {panel.Status}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"AIC:     {panel.AicRating} AIC");
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
          InfoTextBox.AppendText($"ID:         {panelBreaker.Id}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Status:     {panelBreaker.Status}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"AIC:        {panelBreaker.AicRating} AIC");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Amp Rating: {panelBreaker.AmpRating}A");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Poles:      {panelBreaker.NumPoles}P");
          break;
        case NodeType.Disconnect:
          ElectricalEntity.Disconnect disconnect = (ElectricalEntity.Disconnect)entity;
          InfoTextBox.AppendText($"ID:     {disconnect.Id}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Status: {disconnect.Status}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"AIC:    {disconnect.AicRating} AIC");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"AS:     {disconnect.AsSize}AS");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"AF:     {disconnect.AfSize}AF");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Poles:  {disconnect.NumPoles}P");
          break;
        case NodeType.Transformer:
          ElectricalEntity.Transformer transformer = (ElectricalEntity.Transformer)entity;
          InfoTextBox.AppendText($"ID:      {transformer.Id}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Status:  {transformer.Status}");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"AIC:     {transformer.AicRating} AIC");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"KVA:     {transformer.Kva} KVA");
          InfoTextBox.AppendText(Environment.NewLine);
          InfoTextBox.AppendText($"Voltage: {transformer.Voltage}");
          break;
      }
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
      //singleLineNodeTree.SaveAicRatings();
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

    private List<ElectricalEntity.ElectricalEntity> MakeEntitiesFromDistributionBus(
      string distributionBusId,
      Point3d currentPoint
    )
    {
      List<ElectricalEntity.ElectricalEntity> distributionBreakers =
        new List<ElectricalEntity.ElectricalEntity>();
      TreeNode[] nodes = SingleLineTreeView.Nodes.Find(distributionBusId, true);
      if (nodes == null || nodes.Length == 0)
      {
        return distributionBreakers;
      }
      TreeNode thisNode = nodes[0];
      ElectricalEntity.ElectricalEntity thisEntity = (ElectricalEntity.ElectricalEntity)
        thisNode.Tag;

      if (thisEntity.NodeType == NodeType.Meter)
      {
        // Make meter and then breaker
        // add breaker to distributionBreakers
        //
      }
      if (thisEntity.NodeType == NodeType.DistributionBreaker)
      {
        // Just make breaker
        // add breaker to distributionBreakers
      }
      return distributionBreakers;
    }

    private void MakeFieldEntity(
      ElectricalEntity.ElectricalEntity parentEntity,
      Point3d currentPoint
    )
    {
      TreeNode[] nodes = SingleLineTreeView.Nodes.Find(parentEntity.Id, true);
      if (nodes == null || nodes.Length == 0)
      {
        return;
      }
      TreeNode parentNode = nodes[0];
      TreeNode childNode = parentNode.Nodes[0];
      ElectricalEntity.ElectricalEntity childEntity = (ElectricalEntity.ElectricalEntity)
        childNode.Tag;
      if (childEntity.NodeType == NodeType.Panel)
      {
        // Make panel
        // check for PanelBreakers and make (including conduits), then call MakeFieldEntity with next entity and new point for all children from the breakers
      }
      if (childEntity.NodeType == NodeType.Disconnect)
      {
        // Make Disconnect
        // advance currentPoint
        // call MakeFieldEntity(childEntity, currentPoint)
      }
      if (childEntity.NodeType == NodeType.Transformer)
      {
        // Make Transformer
        // advance currentPoint
        // call MakeFieldEntity(childEntity, currentPoint)
      }
    }

    private void MakeDistributionSection(string groupId, Point3d currentPoint)
    {
      currentPoint = new Point3d(currentPoint.X, currentPoint.Y - 0.25, currentPoint.Z);
      string distributionBusId = String.Empty;
      foreach (string groupMember in groupDict[groupId])
      {
        if (distributionBusList.Select(entity => entity.Id).ToArray().Contains(groupMember))
        {
          distributionBusId = groupMember;
        }
      }
      if (String.IsNullOrEmpty(distributionBusId))
      {
        return;
      }

      double totalBusBarWidth = 0;
      TreeNode distributionBusNode = SingleLineTreeView.Nodes.Find(distributionBusId, true)[0];
      foreach (TreeNode distributionBusChild in distributionBusNode.Nodes)
      {
        NodeType nodeType = NodeType.DistributionBreaker;
        if (meterList.Select(entity => entity.Id).ToArray().Contains(distributionBusChild.Name))
        {
          nodeType = NodeType.Meter;
        }
        double width = AggregateEntityWidth(distributionBusChild.Name, nodeType);
        totalBusBarWidth += width;
        // move to center of width
        currentPoint = new Point3d(currentPoint.X + (width / 2), currentPoint.Y, currentPoint.Z);
        List<ElectricalEntity.ElectricalEntity> distributionBreakers =
          MakeEntitiesFromDistributionBus(distributionBusId, currentPoint);
        foreach (ElectricalEntity.ElectricalEntity breaker in distributionBreakers)
        {
          MakeFieldEntity(breaker, currentPoint);
        }
        // move to end of width
        currentPoint = new Point3d(currentPoint.X + (width / 2), currentPoint.Y, currentPoint.Z);
      }
      // Make bus bar based on totalBusBarWidth
    }

    private void MakeGroups(Point3d startingPoint)
    {
      Point3d currentPoint = startingPoint;
      int index = 1;
      foreach (string groupId in groupDict.Keys)
      {
        if (String.IsNullOrEmpty(groupId))
          continue;
        GroupType groupType = InferGroupType(groupDict[groupId]);

        double groupWidth = 2;
        if (groupType == GroupType.MultimeterSection || groupType == GroupType.DistributionSection)
        {
          groupWidth = AggregateGroupWidth(groupId);
          // MakeDistributionSection(groupId, currentPoint)
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
  }
}
