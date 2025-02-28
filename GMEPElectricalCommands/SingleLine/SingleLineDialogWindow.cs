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
using ElectricalCommands.ElectricalEntity;
using GMEPElectricalCommands.GmepDatabase;

namespace ElectricalCommands.SingleLine
{
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
    private Dictionary<string, string> groupDict;
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
      groupDict = new Dictionary<string, string>();
      serviceList.ForEach(service => groupDict.Add(service.Id, GetGroupAssociation(service)));
      meterList.ForEach(meter => groupDict.Add(meter.Id, GetGroupAssociation(meter)));
      mainBreakerList.ForEach(mainBreaker =>
        groupDict.Add(mainBreaker.Id, GetGroupAssociation(mainBreaker))
      );
      distributionBusList.ForEach(distributionBus =>
        groupDict.Add(distributionBus.Id, GetGroupAssociation(distributionBus))
      );
      distributionBreakerList.ForEach(distributionBreaker =>
        groupDict.Add(distributionBreaker.Id, GetGroupAssociation(distributionBreaker))
      );
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
            "Ground Fault: " + (mainBreaker.HasGroundFaultProtection ? "Yes" : "No") // HERE test
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
        //singleLineNodeTree = MakeSingleLineNodeTree();
        //singleLineNodeTree.AggregateWidths();
        //singleLineNodeTree.SetChildStartingPoints(startingPoint);
        //singleLineNodeTree.Make();
      }
      //singleLineNodeTree.SaveAicRatings();
    }
  }
}
