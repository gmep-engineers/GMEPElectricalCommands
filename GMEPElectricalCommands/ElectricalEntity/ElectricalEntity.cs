using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ElectricalCommands.SingleLine;

namespace ElectricalCommands.ElectricalEntity
{
  public class ElectricalEntity
  {
    public string Id,
      Name,
      NodeId;
    public string Status;
    public double AicRating;
    public NodeType NodeType;
    public Point NodePosition;
  }

  public class DistributionBreaker : ElectricalEntity
  {
    public int AmpRating;
    public int NumPoles;
    public bool IsFuseOnly;

    public DistributionBreaker(
      string Id,
      string NodeId,
      string Status,
      int AmpRating,
      int NumPoles,
      bool IsFuseOnly,
      double AicRating,
      System.Drawing.Point NodePosition
    )
    {
      this.Id = Id;
      this.NodeId = NodeId;
      this.Status = Status;
      this.AmpRating = AmpRating;
      this.NumPoles = NumPoles;
      this.IsFuseOnly = IsFuseOnly;
      this.AicRating = AicRating;
      Name = $"{AmpRating}A/{NumPoles}P Distrib. Breaker";
      NodeType = NodeType.DistributionBreaker;
      this.NodePosition = NodePosition;
    }
  }

  public class DistributionBus : ElectricalEntity
  {
    public int AmpRating;

    public DistributionBus(
      string Id,
      string NodeId,
      string Status,
      int AmpRating,
      double AicRating,
      Point NodePosition
    )
    {
      this.Id = Id;
      this.NodeId = NodeId;
      this.Status = Status;
      this.AmpRating = AmpRating;
      this.AicRating = AicRating;
      Name = $"{AmpRating}A Distrib. Bus";
      NodeType = NodeType.DistributionBus;
      this.NodePosition = NodePosition;
    }
  }

  public class MainBreaker : ElectricalEntity
  {
    public int AmpRating;
    public bool HasGroundFaultProtection;
    public bool HasSurgeProtection;
    public int NumPoles;

    public MainBreaker(
      string Id,
      string NodeId,
      string Status,
      int AmpRating,
      bool HasGroundFaultProtection,
      bool HasSurgeProtection,
      int NumPoles,
      double AicRating,
      Point NodePosition
    )
    {
      this.Id = Id;
      this.NodeId = NodeId;
      this.Status = Status;
      this.AmpRating = AmpRating;
      this.HasGroundFaultProtection = HasGroundFaultProtection;
      this.HasSurgeProtection = HasSurgeProtection;
      this.NumPoles = NumPoles;
      this.AicRating = AicRating;
      Name = $"{AmpRating}A/{NumPoles}P Main Breaker";
      NodeType = NodeType.MainBreaker;
      this.NodePosition = NodePosition;
    }
  }

  public class Meter : ElectricalEntity
  {
    public bool HasCts;

    public Meter(
      string Id,
      string NodeId,
      string Status,
      bool HasCts,
      double AicRating,
      Point NodePosition
    )
    {
      this.Id = Id;
      this.NodeId = NodeId;
      this.Status = Status;
      this.HasCts = HasCts;
      this.AicRating = AicRating;
      Name = HasCts ? "CTS Meter" : "Meter";
      NodeType = NodeType.Meter;
      this.NodePosition = NodePosition;
    }
  }

  public class PanelBreaker : ElectricalEntity
  {
    public int AmpRating;
    public int NumPoles;
    public int CircuitNo;

    public PanelBreaker(
      string Id,
      string NodeId,
      string Status,
      int AmpRating,
      int NumPoles,
      int CircuitNo,
      double AicRating,
      Point NodePosition
    )
    {
      this.Id = Id;
      this.NodeId = NodeId;
      this.Status = Status;
      this.AmpRating = AmpRating;
      this.NumPoles = NumPoles;
      this.CircuitNo = CircuitNo;
      this.AicRating = AicRating;
      Name = $"{AmpRating}A/{NumPoles}P Panel Breaker";
      NodeType = NodeType.PanelBreaker;
      this.NodePosition = NodePosition;
    }
  }

  public class NodeLink
  {
    public string Id;
    public string InputConnectorNodeId;
    public string OutputConnectorNodeId;

    public NodeLink(string id, string inputConnectorNodeId, string outputConnectorNodeId)
    {
      Id = id;
      InputConnectorNodeId = inputConnectorNodeId;
      OutputConnectorNodeId = outputConnectorNodeId;
    }
  }

  public class GroupNode
  {
    public string Id;
    public int Width;
    public int Height;
    public string Name;
    public Point NodePosition;
    public string Status;

    public GroupNode(
      string id,
      int width,
      int height,
      string name,
      Point nodePosition,
      string status
    )
    {
      Id = id;
      Width = width;
      Height = height;
      Name = name;
      NodePosition = nodePosition;
      Status = status;
    }
  }
}
