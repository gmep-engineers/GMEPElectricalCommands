using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands
{
  public class SpoonShapeCommands
  {
    [CommandMethod("SP")]
    public void SP()
    {
      double scale = 12;
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Editor ed = doc.Editor;
      if (
        CADObjectCommands.Scale <= 0
        && (CADObjectCommands.IsInModel() || CADObjectCommands.IsInLayoutViewport())
      )
      {
        CADObjectCommands.SetScale();
        if (CADObjectCommands.Scale <= 0)
          return;
      }
      if (CADObjectCommands.IsInModel() || CADObjectCommands.IsInLayoutViewport())
      {
        scale = CADObjectCommands.Scale;
      }

      PromptPointOptions ppo = new PromptPointOptions("\nSelect start point:");
      PromptPointResult ppr = ed.GetPoint(ppo);
      if (ppr.Status != PromptStatus.OK)
        return;

      Point3d firstClickPoint = ppr.Value;
      SpoonJig jig = new SpoonJig(firstClickPoint, scale);

      PromptResult res = ed.Drag(jig);
      if (res.Status != PromptStatus.OK)
        return;
      Vector3d direction = jig.endPoint - firstClickPoint;
      double angle = direction.GetAngleTo(Vector3d.XAxis, Vector3d.ZAxis);

      Point3d secondClickPoint = jig.endPoint;
      Point3d thirdClickPoint = Point3d.Origin;
      bool thirdClickOccurred = false;

      using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
      {
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite);
        btr.AppendEntity(jig.line);
        tr.AddNewlyCreatedDBObject(jig.line, true);

        btr.AppendEntity(jig.arc);
        tr.AddNewlyCreatedDBObject(jig.arc, true);

        tr.Commit();
      }

      if (angle != 0 && angle != Math.PI)
      {
        DynamicLineJig lineJig = new DynamicLineJig(jig.endPoint, scale);
        res = ed.Drag(lineJig);
        if (res.Status == PromptStatus.OK)
        {
          using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
          {
            BlockTableRecord btr = (BlockTableRecord)
              tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite);
            btr.AppendEntity(lineJig.line);
            tr.AddNewlyCreatedDBObject(lineJig.line, true);

            thirdClickPoint = lineJig.line.EndPoint;
            thirdClickOccurred = true;

            tr.Commit();
          }
        }
      }

      Point3d textAlignmentReferencePoint = thirdClickOccurred ? thirdClickPoint : secondClickPoint;
      Point3d comparisonPoint = thirdClickOccurred ? secondClickPoint : firstClickPoint;

      using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
      {
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite);

        // Create MText object with dynamic text height based on scale
        double textHeight = 1.125 / scale;
        MText mText = new MText();
        mText.SetDatabaseDefaults();
        mText.TextHeight = textHeight;

        // Retrieve the "rpm" text style
        using (Transaction tr2 = doc.Database.TransactionManager.StartTransaction())
        {
          TextStyleTable textStyleTable = (TextStyleTable)
            tr2.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead);
          if (textStyleTable.Has("rpm"))
          {
            mText.TextStyleId = textStyleTable["rpm"];
          }
          else
          {
            ed.WriteMessage("\nText style 'rpm' not found. Using default text style.");
            mText.TextStyleId = doc.Database.Textstyle;
          }
          tr2.Commit();
        }

        mText.Width = 0;
        mText.Layer = "E-TXT1";

        // Determine justification and position
        if (textAlignmentReferencePoint.X > comparisonPoint.X)
        {
          mText.Attachment = AttachmentPoint.TopLeft;
          mText.Location = new Point3d(
            textAlignmentReferencePoint.X + 0.25 / scale,
            textAlignmentReferencePoint.Y + textHeight / 2,
            textAlignmentReferencePoint.Z
          );
        }
        else
        {
          mText.Attachment = AttachmentPoint.TopRight;
          mText.Location = new Point3d(
            textAlignmentReferencePoint.X - 0.25 / scale,
            textAlignmentReferencePoint.Y + textHeight / 2,
            textAlignmentReferencePoint.Z
          );
        }

        btr.AppendEntity(mText);
        tr.AddNewlyCreatedDBObject(mText, true);

        tr.Commit();

        doc.Editor.Command("_.MTEDIT", mText.ObjectId);
      }
    }
  }

  public class DynamicLineJig : DrawJig
  {
    private Point3d startPoint;
    private Point3d endPoint;
    private double scale;
    public Line line;

    public DynamicLineJig(Point3d startPt, double scale)
    {
      this.scale = scale;
      startPoint = startPt;
      endPoint = startPt;
      line = new Line(startPoint, startPoint);
      line.Layer = "E-TXT1";
    }

    protected override bool WorldDraw(WorldDraw draw)
    {
      if (line != null)
      {
        draw.Geometry.Draw(line);
      }
      return true;
    }

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
      JigPromptPointOptions opts = new JigPromptPointOptions("\nSelect end point:");
      opts.BasePoint = startPoint;
      opts.UseBasePoint = true;
      opts.Cursor = CursorType.RubberBand;

      PromptPointResult res = prompts.AcquirePoint(opts);
      if (res.Status != PromptStatus.OK)
        return SamplerStatus.Cancel;

      if (endPoint.DistanceTo(res.Value) < Tolerance.Global.EqualPoint)
        return SamplerStatus.NoChange;

      endPoint = new Point3d(res.Value.X, startPoint.Y, startPoint.Z);
      line.EndPoint = endPoint;

      return SamplerStatus.OK;
    }
  }

  public class SpoonJig : DrawJig
  {
    private Point3d firstClickPoint;
    private Point3d startPoint;
    public Point3d endPoint { get; private set; }
    private Point3d rotationPoint;
    private double scale;
    public Line line;
    public Arc arc;

    public SpoonJig(Point3d firstClick, double scale)
    {
      firstClickPoint = firstClick;
      rotationPoint = firstClickPoint;
      this.scale = scale;
      startPoint = rotationPoint + new Vector3d(-3 * (0.25 / scale), 0, 0);
      endPoint = startPoint;
      line = new Line(startPoint, startPoint);
      line.Layer = "E-TXT1";
      arc = new Arc();
      arc.Layer = "E-TXT1";
    }

    public PromptStatus DragMe()
    {
      Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
      PromptResult res;
      do
      {
        res = ed.Drag(this);
      } while (res.Status == PromptStatus.Other);
      return res.Status;
    }

    protected override bool WorldDraw(WorldDraw draw)
    {
      if (line != null)
      {
        draw.Geometry.Draw(line);
      }
      if (arc != null && arc.StartAngle != arc.EndAngle)
      {
        draw.Geometry.Draw(arc);
      }
      return true;
    }

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
      JigPromptPointOptions opts = new JigPromptPointOptions("\nSelect end point:");
      opts.BasePoint = rotationPoint;
      opts.UseBasePoint = true;
      opts.Cursor = CursorType.RubberBand;

      PromptPointResult res = prompts.AcquirePoint(opts);
      if (res.Status != PromptStatus.OK)
        return SamplerStatus.Cancel;

      if (endPoint.DistanceTo(res.Value) < Tolerance.Global.EqualPoint)
        return SamplerStatus.NoChange;

      endPoint = res.Value;

      Vector3d direction = (endPoint - rotationPoint).GetNormal();
      startPoint = rotationPoint + direction * -3 * (0.25 / scale);

      UpdateGeometry();
      return SamplerStatus.OK;
    }

    private void UpdateGeometry()
    {
      line.StartPoint = startPoint;
      line.EndPoint = endPoint;

      Vector3d direction = (endPoint - startPoint).GetNormal();
      Vector3d perpendicular = new Vector3d(-direction.Y, direction.X, 0);
      Point3d secondPoint =
        startPoint + direction * 2 * (0.25 / scale) + perpendicular * 4 * (0.25 / scale);
      Point3d thirdPoint = startPoint + direction * 6 * (0.25 / scale);

      if ((endPoint - startPoint).Length > 6 * (0.25 / scale))
      {
        arc.SetDatabaseDefaults();
        arc.Center = Arc3PCenter(startPoint, secondPoint, thirdPoint);
        arc.Radius = (startPoint - arc.Center).Length;

        Vector3d startVector = startPoint - arc.Center;
        Vector3d endVector = thirdPoint - arc.Center;

        arc.StartAngle = Math.Atan2(startVector.Y, startVector.X);
        arc.EndAngle = Math.Atan2(endVector.Y, endVector.X);
      }
    }

    private Point3d Arc3PCenter(Point3d p1, Point3d p2, Point3d p3)
    {
      CircularArc3d tempArc = new CircularArc3d(p1, p2, p3);
      return tempArc.Center;
    }
  }

  public class ArrowJig : EntityJig
  {
    private Point3d _insertionPoint;
    private Point3d _panelLocation;
    private Vector3d _direction;

    public ArrowJig(BlockReference blockRef, Point3d panelLocation)
      : base(blockRef)
    {
      _insertionPoint = Point3d.Origin;
      _panelLocation = panelLocation;
    }

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
      JigPromptPointOptions pointOptions = new JigPromptPointOptions("\nSpecify insertion point: ");
      PromptPointResult pointResult = prompts.AcquirePoint(pointOptions);

      if (pointResult.Status == PromptStatus.OK)
      {
        if (_insertionPoint == pointResult.Value)
          return SamplerStatus.NoChange;

        _insertionPoint = pointResult.Value;
        _direction = _panelLocation - _insertionPoint;
        return SamplerStatus.OK;
      }

      return SamplerStatus.Cancel;
    }

    protected override bool Update()
    {
      ((BlockReference)Entity).Position = _insertionPoint;
      ((BlockReference)Entity).Rotation = Math.Atan2(_direction.Y, _direction.X) - Math.PI / 2;
      return true;
    }

    public Point3d InsertionPoint => _insertionPoint;
  }

  public class LabelJig : DrawJig
  {
    private Point3d firstClickPoint;
    private Point3d startPoint;
    public Point3d endPoint { get; private set; }
    private Point3d rotationPoint;
    private double scale;
    public Line line;

    public LabelJig(Point3d firstClick, double scale)
    {
      firstClickPoint = firstClick;
      rotationPoint = firstClickPoint;
      this.scale = scale;
      startPoint = rotationPoint;
      endPoint = startPoint;
      line = new Line(startPoint, startPoint);
      line.Layer = "E-TXT1";
    }

    public PromptStatus DragMe()
    {
      Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
      PromptResult res;
      do
      {
        res = ed.Drag(this);
      } while (res.Status == PromptStatus.Other);
      return res.Status;
    }

    protected override bool WorldDraw(WorldDraw draw)
    {
      if (line != null)
      {
        draw.Geometry.Draw(line);
      }
      return true;
    }

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
      JigPromptPointOptions opts = new JigPromptPointOptions("\nSelect end point:");
      opts.BasePoint = rotationPoint;
      opts.UseBasePoint = true;
      opts.Cursor = CursorType.RubberBand;

      PromptPointResult res = prompts.AcquirePoint(opts);
      if (res.Status != PromptStatus.OK)
        return SamplerStatus.Cancel;

      if (endPoint.DistanceTo(res.Value) < Tolerance.Global.EqualPoint)
        return SamplerStatus.NoChange;

      endPoint = res.Value;

      startPoint = rotationPoint;

      UpdateGeometry();
      return SamplerStatus.OK;
    }

    private void UpdateGeometry()
    {
      line.StartPoint = startPoint;
      line.EndPoint = endPoint;

      Vector3d direction = (endPoint - startPoint).GetNormal();
      Vector3d perpendicular = new Vector3d(-direction.Y, direction.X, 0);
      Point3d secondPoint =
        startPoint + direction * 2 * (0.25 / scale) + perpendicular * 4 * (0.25 / scale);
      Point3d thirdPoint = startPoint + direction * 6 * (0.25 / scale);
    }
  }
}
