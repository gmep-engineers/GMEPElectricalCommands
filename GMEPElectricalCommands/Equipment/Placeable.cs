using System;
using System.Windows.Forms;
using Autodesk.AutoCAD.Geometry;

namespace ElectricalCommands.Equipment
{
  public class Placeable
  {
    public string id,
      name,
      parentId,
      parentName;
    public int parentDistance;
    public double aicRating;
    public Point3d loc;
  }

  public class LightingControl : Placeable
  {
    public string controlType;
    public bool occupancy;

    public LightingControl(string id, string name, string controlType, int occupancy)
    {
      this.id = id;
      this.name = name;
      this.controlType = controlType;
      this.occupancy = occupancy != 0;
    }
  }

  public class LightingFixture : Placeable
  {
    public int voltage,
      qty;
    public double wattage,
      paperSpaceScale,
      labelTransformHX,
      labelTransformHY,
      labelTransformVX,
      labelTransformVY;
    public string blockName,
      controlId,
      description,
      mounting,
      manufacturer,
      modelNo,
      notes;
    public bool rotate,
      emCapable;

    public LightingFixture(
      string id,
      string parentId,
      string parentName,
      string name,
      string controlId,
      string blockName,
      int voltage,
      double wattage,
      string description,
      int qty,
      string mounting,
      string manufacturer,
      string modelNo,
      string notes,
      int rotate,
      double paperSpaceScale,
      int emCapable,
      double labelTransformHX,
      double labelTransformHY,
      double labelTransformVX,
      double labelTransformVY
    )
    {
      this.id = id;
      this.parentId = parentId;
      this.parentName = parentName;
      this.name = name;
      this.blockName = blockName;
      this.voltage = voltage;
      this.wattage = wattage;
      this.controlId = controlId;
      this.description = description;
      this.qty = qty;
      this.mounting = mounting;
      this.manufacturer = manufacturer;
      this.modelNo = modelNo;
      this.notes = notes;
      this.rotate = rotate != 0;
      this.paperSpaceScale = paperSpaceScale;
      this.emCapable = emCapable != 0;
      this.labelTransformHX = labelTransformHX;
      this.labelTransformHY = labelTransformHY;
      this.labelTransformVX = labelTransformVX;
      this.labelTransformVY = labelTransformVY;
    }
  }

  public class Equipment : Placeable
  {
    public string description,
      category,
      hp;
    public int voltage,
      mca,
      mountingHeight,
      circuit;
    public double fla;
    public bool is3Phase;

    public Equipment(
      string id = "",
      string parentId = "",
      string parentName = "",
      string name = "",
      string description = "",
      string category = "",
      int voltage = 0,
      double fla = 0,
      bool is3Phase = false,
      int parentDistance = -1,
      double xLoc = 0,
      double yLoc = 0,
      int mca = -1,
      string hp = "",
      int mountingHeight = 18,
      int circuit = 0
    )
    {
      this.id = id;
      this.parentId = parentId;
      this.parentName = parentName;
      this.name = name.ToUpper();
      this.description = description;
      this.category = category;
      this.voltage = voltage;
      this.fla = fla;
      this.is3Phase = is3Phase;
      loc = new Point3d(xLoc, yLoc, 0);
      this.parentDistance = parentDistance;
      this.mca = mca;
      this.hp = hp;
      this.mountingHeight = mountingHeight;
      this.circuit = circuit;
    }
  }

  public class Panel : Placeable
  {
    public bool isDistribution;
    public bool isMultiMeter;
    public int busSize;
    public string voltage;

    public Panel(
      string _id,
      string pId,
      string n,
      int pDist = -1,
      double xLoc = 0,
      double yLoc = 0,
      int isDistrib = 0,
      int isMm = 0,
      int bus = 0,
      string volt = "",
      double aicRating = 0
    )
    {
      id = _id;
      parentId = pId;
      name = n;
      parentDistance = pDist;
      loc = new Point3d(xLoc, yLoc, 0);
      isDistribution = isDistrib == 1;
      isMultiMeter = isMm == 1;
      busSize = bus;
      voltage = volt;
      this.aicRating = aicRating;
    }
  }

  public class Transformer : Placeable
  {
    public double kva;
    public string voltageSpec;

    public Transformer(
      string _id,
      string pId,
      string n,
      int pDist = -1,
      double xLoc = 0,
      double yLoc = 0,
      double kva = 0,
      string volt = "",
      double aicRating = 0
    )
    {
      id = _id;
      parentId = pId;
      name = n;
      parentDistance = pDist;
      loc = new Point3d(xLoc, yLoc, 0);
      this.kva = kva;
      voltageSpec = volt;
      this.aicRating = aicRating;
    }
  }
}
