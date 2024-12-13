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
    public Point3d loc;
  }

  public class Equipment : Placeable
  {
    public string description,
      category;
    public int voltage;
    public bool is3Phase;

    public Equipment(
      string _id = "",
      string pId = "",
      string pName = "",
      string eqNo = "",
      string desc = "",
      string cat = "",
      int volts = 0,
      bool is3Ph = false,
      int pDist = -1,
      double xLoc = 0,
      double yLoc = 0
    )
    {
      id = _id;
      parentId = pId;
      parentName = pName;
      name = eqNo.ToUpper();
      description = desc;
      category = cat;
      voltage = volts;
      is3Phase = is3Ph;
      loc = new Point3d(xLoc, yLoc, 0);
      parentDistance = pDist;
    }
  }

  public class Panel : Placeable
  {
    public Panel(string _id, string pId, string n, int pDist = -1, double xLoc = 0, double yLoc = 0)
    {
      id = _id;
      parentId = pId;
      name = n;
      parentDistance = pDist;
      loc = new Point3d(xLoc, yLoc, 0);
    }
  }

  public class Transformer : Placeable
  {
    public Transformer(
      string _id,
      string pId,
      string n,
      int pDist = -1,
      double xLoc = 0,
      double yLoc = 0
    )
    {
      id = _id;
      parentId = pId;
      name = n;
      parentDistance = pDist;
      loc = new Point3d(xLoc, yLoc, 0);
    }
  }
}
