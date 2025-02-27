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
      nodeLinkList = gmepDb.GetNodesLinks(projectId);
      SingleLineTreeView.BeginUpdate();
      PopulateTreeView();
      SingleLineTreeView.EndUpdate();
    }

    public void PopulateTreeView()
    {
      foreach (ElectricalEntity.Service service in serviceList)
      {
        TreeNode serviceNode = SingleLineTreeView.Nodes.Add(service.Id, service.Name);
        PopulateFromService(serviceNode, service.NodeId);
      }
    }

    public void PopulateFromService(TreeNode node, string serviceNodeId)
    {
      foreach (ElectricalEntity.Meter meter in meterList)
      {
        Console.WriteLine(meter.NodeId);
        Console.WriteLine(serviceNodeId);
        if (VerifyNodeLink(serviceNodeId, meter.NodeId))
        {
          TreeNode meterNode = node.Nodes.Add(meter.Id, meter.Name);
          PopulateFromMainMeter(meterNode, meter.NodeId);
        }
      }
      foreach (ElectricalEntity.MainBreaker mainBreaker in mainBreakerList)
      {
        if (VerifyNodeLink(serviceNodeId, mainBreaker.NodeId))
        {
          TreeNode mainBreakerNode = node.Nodes.Add(mainBreaker.Id, mainBreaker.Name);
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
          PopulateFromDistributionBreaker(distributionBreakerNode, distributionBreaker.NodeId);
        }
      }
    }

    public void PopulateFromDistributionBreaker(TreeNode node, string distributionBreakerNodeId)
    {
      foreach (ElectricalEntity.Panel panel in panelList)
      {
        if (VerifyNodeLink(distributionBreakerNodeId, panel.Id))
        {
          TreeNode panelNode = node.Nodes.Add(panel.Id, panel.Name);
          PopulateFromPanel(panelNode, panel.NodeId);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(distributionBreakerNodeId, disconnect.NodeId))
        {
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          PopulateFromDisconnect(disconnectNode, disconnect.NodeId);
        }
      }
      foreach (ElectricalEntity.Transformer transformer in transformerList)
      {
        if (VerifyNodeLink(distributionBreakerNodeId, transformer.NodeId))
        {
          TreeNode transformerNode = node.Nodes.Add(transformer.Id, transformer.Name);
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
          PopulateFromPanelBreaker(panelBreakerNode, panelBreaker.NodeId);
        }
      }
      foreach (ElectricalEntity.Panel panel in panelList)
      {
        if (VerifyNodeLink(panelNodeId, panel.Id))
        {
          TreeNode panelNode = node.Nodes.Add(panel.Id, panel.Name);
          PopulateFromPanel(panelNode, panel.NodeId);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(panelNodeId, disconnect.NodeId))
        {
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          PopulateFromDisconnect(disconnectNode, disconnect.NodeId);
        }
      }
      foreach (ElectricalEntity.Transformer transformer in transformerList)
      {
        if (VerifyNodeLink(panelNodeId, transformer.NodeId))
        {
          TreeNode transformerNode = node.Nodes.Add(transformer.Id, transformer.Name);
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
          PopulateFromPanel(panelNode, panel.NodeId);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(panelBreakerNodeId, disconnect.NodeId))
        {
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          PopulateFromDisconnect(disconnectNode, disconnect.NodeId);
        }
      }
      foreach (ElectricalEntity.Transformer transformer in transformerList)
      {
        if (VerifyNodeLink(panelBreakerNodeId, transformer.NodeId))
        {
          TreeNode transformerNode = node.Nodes.Add(transformer.Id, transformer.Name);
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
          PopulateFromPanel(panelNode, panel.NodeId);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(disconnectNodeId, disconnect.NodeId))
        {
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          PopulateFromDisconnect(disconnectNode, disconnect.NodeId);
        }
      }
      foreach (ElectricalEntity.Transformer transformer in transformerList)
      {
        if (VerifyNodeLink(disconnectNodeId, transformer.NodeId))
        {
          TreeNode transformerNode = node.Nodes.Add(transformer.Id, transformer.Name);
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
          PopulateFromPanel(panelNode, panel.NodeId);
        }
      }
      foreach (ElectricalEntity.Disconnect disconnect in disconnectList)
      {
        if (VerifyNodeLink(transformerNodeId, disconnect.NodeId))
        {
          TreeNode disconnectNode = node.Nodes.Add(disconnect.Id, disconnect.Name);
          PopulateFromDisconnect(disconnectNode, disconnect.NodeId);
        }
      }
    }

    public bool VerifyNodeLink(string outputConnectorNodeId, string inputConnectorNodeId)
    {
      foreach (ElectricalEntity.NodeLink link in nodeLinkList)
      {
        Console.WriteLine("nodeA " + link.OutputConnectorNodeId);
        Console.WriteLine("nodeB " + link.InputConnectorNodeId);
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
