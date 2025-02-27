using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Autodesk.AutoCAD.Geometry;

namespace ElectricalCommands.ElectricalEntity
{
  public class PlaceableElectricalEntity : ElectricalEntity
  {
    public string ParentId;
    public string ParentName;
    public int ParentDistance;
    public Point3d Location;
    public bool IsHidden;
  }

  public class LightingControl : PlaceableElectricalEntity
  {
    public string ControlType;
    public bool HasOccupancy;

    public LightingControl(string Id, string Name, string ControlTypeId, bool HasOccupancy)
    {
      this.Id = Id;
      this.Name = Name;
      this.ControlType = ControlTypeId;
      this.HasOccupancy = HasOccupancy;
    }
  }

  public class LightingFixture : PlaceableElectricalEntity
  {
    public int Voltage,
      Qty;
    public double Wattage,
      PaperSpaceScale,
      LabelTransformHX,
      LabelTransformHY,
      LabelTransformVX,
      LabelTransformVY;
    public string BlockName,
      ControlId,
      Description,
      Mounting,
      Manufacturer,
      ModelNo,
      Notes;
    public bool Rotate,
      EmCapable;

    public LightingFixture(
      string Id,
      string ParentId,
      string ParentName,
      string Name,
      string ControlId,
      string BlockName,
      int Voltage,
      double Wattage,
      string Description,
      int Qty,
      string Mounting,
      string Manufacturer,
      string ModelNo,
      string Notes,
      bool Rotate,
      double PaperSpaceScale,
      bool EmCapable,
      double LabelTransformHX,
      double LabelTransformHY,
      double LabelTransformVX,
      double LabelTransformVY
    )
    {
      this.Id = Id;
      this.ParentId = ParentId;
      this.ParentName = ParentName;
      this.Name = Name;
      this.BlockName = BlockName;
      this.Voltage = Voltage;
      this.Wattage = Wattage;
      this.ControlId = ControlId;
      this.Description = Description;
      this.Qty = Qty;
      this.Mounting = Mounting;
      this.Manufacturer = Manufacturer;
      this.ModelNo = ModelNo;
      this.Notes = Notes;
      this.Rotate = Rotate;
      this.PaperSpaceScale = PaperSpaceScale;
      this.EmCapable = EmCapable;
      this.LabelTransformHX = LabelTransformHX;
      this.LabelTransformHY = LabelTransformHY;
      this.LabelTransformVX = LabelTransformVX;
      this.LabelTransformVY = LabelTransformVY;
    }
  }

  public class DistributionSection : PlaceableElectricalEntity
  {
    public DistributionBus DistributionBus;
    public List<Meter> Meters;
    public List<DistributionBreaker> Breakers;

    public DistributionSection() { }
  }

  public class Service : PlaceableElectricalEntity
  {
    public int AmpRating;
    public string Voltage;
    public List<Meter> Meters;
    public List<MainBreaker> MainBreakers;
    public DistributionSection DistributionSection;

    public Service(
      string Id,
      string Name,
      string Status,
      int AmpRating,
      string Voltage,
      double AicRating
    )
    {
      this.Id = Id;
      this.Name = Name;
      this.Status = Status;
      this.AmpRating = AmpRating;
      this.Voltage = Voltage;
      this.AicRating = AicRating;
    }
  }

  public class Equipment : PlaceableElectricalEntity
  {
    public string Description,
      Hp,
      Category;
    public int Voltage,
      MountingHeight,
      Circuit;
    public double Fla,
      Mca;
    public bool Is3Phase,
      HasPlug;

    public Equipment(
      string Id,
      string ParentId,
      string ParentName,
      string Name,
      string Description,
      string Category,
      int Voltage,
      double Fla,
      bool Is3Phase,
      int ParentDistance,
      double LocationX,
      double LocationY,
      float Mca,
      string Hp,
      int MountingHeight,
      int Circuit,
      bool HasPlug,
      bool Hidden
    )
    {
      this.Id = Id;
      this.ParentId = ParentId;
      this.ParentName = ParentName;
      this.Name = Name.ToUpper();
      this.Description = Description;
      this.Category = Category;
      this.Voltage = Voltage;
      this.Fla = Fla;
      this.Is3Phase = Is3Phase;
      this.Location = new Point3d(LocationX, LocationY, 0);
      this.ParentDistance = ParentDistance;
      this.Mca = Mca;
      this.Hp = Hp;
      this.MountingHeight = MountingHeight;
      this.Circuit = Circuit;
      this.HasPlug = HasPlug;
      this.IsHidden = Hidden;
    }
  }

  public class Disconnect : PlaceableElectricalEntity
  {
    public int AsSize;
    public int AfSize;
    public int NumPoles;

    public Disconnect(
      string Id,
      string ParentId,
      int ParentDistance,
      string NodeId,
      string Status,
      int AsSize,
      int AfSize,
      int NumPoles,
      double AicRating
    )
    {
      this.Id = Id;
      this.ParentId = ParentId;
      this.ParentDistance = ParentDistance;
      this.NodeId = NodeId;
      this.AsSize = AsSize;
      this.AfSize = AfSize;
      this.NumPoles = NumPoles;
      this.Status = Status;
      this.AicRating = AicRating;
      Name = $"{AsSize}AS/{AfSize}AF/{NumPoles}P Disconnect";
    }
  }

  public class Panel : PlaceableElectricalEntity
  {
    public int BusAmpRating;
    public int MainAmpRating;
    public string Voltage;
    public bool IsMlo;
    public List<PanelBreaker> Breakers;

    public Panel(
      string Id,
      string ParentId,
      string Name,
      int ParentDistance,
      double LocationX,
      double LocationY,
      int BusAmpRating,
      int MainAmpRating,
      string Voltage,
      double AicRating,
      bool IsHidden,
      string NodeId,
      string Status
    )
    {
      this.Id = Id;
      this.ParentId = ParentId;
      this.Name = Name;
      this.ParentDistance = ParentDistance;
      Location = new Point3d(LocationX, LocationY, 0);
      this.BusAmpRating = BusAmpRating;
      this.MainAmpRating = MainAmpRating;
      this.Voltage = Voltage;
      this.AicRating = AicRating;
      this.IsHidden = IsHidden;
      this.NodeId = NodeId;
      this.Status = Status;
    }
  }

  public class Transformer : PlaceableElectricalEntity
  {
    public double Kva;
    public string Voltage;

    public Transformer(
      string Id,
      string ParentId,
      string Name,
      int ParentDistance,
      double LocationX,
      double LocationY,
      double Kva,
      string Voltage,
      double AicRating,
      bool IsHidden,
      string NodeId,
      string Status
    )
    {
      this.Id = Id;
      this.ParentId = ParentId;
      this.Name = Name;
      this.ParentDistance = ParentDistance;
      Location = new Point3d(LocationX, LocationY, 0);
      this.Kva = Kva;
      this.Voltage = Voltage;
      this.AicRating = AicRating;
      this.IsHidden = IsHidden;
      this.NodeId = NodeId;
      this.Status = Status;
    }
  }
}
