using System.Collections.Generic;
using Autodesk.AutoCAD.Geometry;
using DocumentFormat.OpenXml.Drawing.Diagrams;

namespace ElectricalCommands.Equipment
{
  public class SingleLine
  {
    public double width;
    public string type;
    public string name;
    public string id;
    public List<SingleLine> children;
    public Point3d startingPoint;
    public Point3d endingPoint;

    public SingleLine() { }

    public double AggregateWidths()
    {
      double sum = width;
      foreach (SingleLine child in children)
      {
        sum += child.AggregateWidths();
      }
      return sum;
    }

    public virtual void SetChildStartingPoints(Point3d startingPoint)
    {
      this.startingPoint = startingPoint;

      foreach (SingleLine child in children)
      {
        child.SetChildStartingPoints(startingPoint);
      }
    }

    public void SetChildEndingPoint(Point3d endingPoint)
    {
      this.endingPoint = endingPoint;
    }

    public virtual void Make()
    {
      foreach (SingleLine child in children)
      {
        child.Make();
      }
    }
  }

  public class SLServiceFeeder : SingleLine
  {
    public SLServiceFeeder(string id, string name)
    {
      type = "service feeder";
      width = 2.5;
      this.name = name;
      this.id = id;
    }

    public override void SetChildStartingPoints(Point3d startingPoint)
    {
      double offset = 0;
      foreach (SingleLine child in children)
      {
        child.SetChildStartingPoints(
          new Point3d(startingPoint.X + width + offset, startingPoint.Y, startingPoint.Z)
        );
        offset += width + 1;
      }
    }
  }

  public class SLMainBreakerSection : SingleLine
  {
    public int breakerSize;

    public SLMainBreakerSection(string id, string name)
    {
      type = "main breaker section";
      width = 1;
      this.name = name;
      this.id = id;
    }
  }

  public class SLDistributionSection : SingleLine
  {
    public SLDistributionSection(string id, string name)
    {
      type = "distribution section";
      width = 1;
      this.name = name;
      this.id = id;
    }
  }

  public class SLPanel : SingleLine
  {
    public bool isDistribution;

    public SLPanel(string id, string name)
    {
      type = "panel";
      width = 2;
      this.name = name;
      this.id = id;
    }

    public override void SetChildStartingPoints(Point3d startingPoint)
    {
      this.startingPoint = startingPoint;
      if (isDistribution)
      {
        double offset = 0.5;
        foreach (SingleLine child in children)
        {
          child.SetChildEndingPoint(
            new Point3d(
              startingPoint.X + (child.width / 2) + offset,
              startingPoint.Y - 5,
              startingPoint.Z
            )
          );
          child.SetChildStartingPoints(
            new Point3d(
              startingPoint.X + (child.width / 2) + offset,
              startingPoint.Y - 0.25,
              startingPoint.Z
            )
          );
          offset += child.width;
        }
      }
      else
      {
        int index = 0;
        foreach (SingleLine child in children)
        {
          if (index == 0)
          {
            child.SetChildEndingPoint(new Point3d(endingPoint.X + 1, endingPoint.Y - 3.25, 0));
            child.SetChildStartingPoints(
              new Point3d(endingPoint.X + (5 / 16), endingPoint.Y - (7 / 8), 0)
            );
          }
        }
        if (children.Count == 1) { }
        if (children.Count == 2) { }
        if (children.Count == 3) { }
        if (children.Count == 4) { }
      }
    }
  }

  public class SLTransformer : SingleLine
  {
    public SLTransformer(string id, string name)
    {
      type = "transformer";
      width = 0;
      this.name = name;
      this.id = id;
    }
  }

  public class SLDisconnect : SingleLine
  {
    public SLDisconnect(string name)
    {
      type = "disconnect";
      width = 0;
      this.name = name;
      this.id = id;
    }
  }

  public class SLConduit : SingleLine
  {
    bool isVertical = true;
    string conduitSize,
      wireSize,
      distance;

    public SLConduit(
      bool isVertical,
      string conduitSize,
      string wireSize,
      string distance,
      string name = ""
    )
    {
      type = "conduit";
      width = 1;
      this.name = name;
      this.isVertical = isVertical;
      this.conduitSize = conduitSize;
      this.wireSize = wireSize;
      this.distance = distance;
    }
  }

  public class SLPanelBreaker : SingleLine
  {
    bool isVertical = true;
    string breakerSize;
    string numPoles;

    public SLPanelBreaker(bool isVertical, string breakerSize, string numPoles)
    {
      this.isVertical = isVertical;
      this.breakerSize = breakerSize;
      this.numPoles = numPoles;
      width = 0;
    }
  }
}
