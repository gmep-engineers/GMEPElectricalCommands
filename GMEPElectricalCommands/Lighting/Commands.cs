using System;
using System.Buffers;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Accord.MachineLearning;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Dreambuild.AutoCAD;
using ElectricalCommands.ElectricalEntity;
using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.ML;
using Emgu.CV.Structure;
using Emgu.CV.Util;
using GMEPElectricalCommands.GmepDatabase;
using Newtonsoft.Json;
using TriangleNet.Meshing.Algorithm;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Color = System.Drawing.Color;
using Group = Autodesk.AutoCAD.DatabaseServices.Group;
using Panel = ElectricalCommands.ElectricalEntity.Panel;

namespace ElectricalCommands.Lighting
{
  public class Commands
  {
    // update captureclickarea to use the blockdata objects to make a block and create a block reference that sticks to the nearest gray line
    [CommandMethod("StickySymbol")]
    public void StickySymbol()
    {
      int WIDTH = 200;
      int HEIGHT = 200;

      Document doc = Application.DocumentManager.MdiActiveDocument;
      Editor ed = doc.Editor;

      PromptPointResult ppr = ed.GetPoint("\nClick a point: ");
      if (ppr.Status != PromptStatus.OK)
        return;
      Point3d clickPoint = ppr.Value;

      System.Windows.Point point = ed.PointToScreen(clickPoint, 0);

      System.Drawing.Point cursorPosition = System.Windows.Forms.Cursor.Position;

      int centerX = WIDTH / 2;
      int centerY = HEIGHT / 2;

      Rectangle captureRectangle = new Rectangle(
        cursorPosition.X - centerX,
        cursorPosition.Y - centerY,
        WIDTH,
        HEIGHT
      );

      using (Bitmap bitmap = new Bitmap(captureRectangle.Width, captureRectangle.Height))
      {
        using (Graphics g = Graphics.FromImage(bitmap))
        {
          g.CopyFromScreen(captureRectangle.Location, Point.Empty, captureRectangle.Size);
        }

        // Create a new bitmap to hold the circular image
        Bitmap circularBitmap = new Bitmap(bitmap.Width, bitmap.Height, bitmap.PixelFormat);

        using (Graphics g = Graphics.FromImage(circularBitmap))
        {
          // Create a circular path that fits within the bounds of the bitmap
          using (GraphicsPath path = new GraphicsPath())
          {
            path.AddEllipse(0, 0, bitmap.Width, bitmap.Height);
            g.SetClip(path);

            // Draw the original bitmap onto the graphics of the new bitmap
            g.DrawImage(bitmap, 0, 0);
          }
        }

        // Step 1: Get a list of every gray pixel
        List<System.Drawing.Point> grayPixels = new List<System.Drawing.Point>();
        for (int x = 0; x < bitmap.Width; x++)
        {
          for (int y = 0; y < bitmap.Height; y++)
          {
            Color pixelColor = bitmap.GetPixel(x, y);
            if (pixelColor.R == pixelColor.G && pixelColor.G == pixelColor.B && pixelColor.R != 0)
            {
              grayPixels.Add(new System.Drawing.Point(x, y));
            }
          }
        }

        var closestGrayPixel = FindClosestGrayPixelWithBlackBorder(
          grayPixels,
          bitmap,
          centerX,
          centerY
        );
        var closestGrayPixelsInPoint = FindGrayPixelsAroundClosest(
          grayPixels,
          bitmap,
          centerX,
          centerY,
          closestGrayPixel
        );
        var closestGrayPixelsInLine = FindClosestGrayPixelsWithBlackBorders(
          grayPixels,
          bitmap,
          30,
          centerX,
          centerY
        );

        Tuple<List<double>, List<double>> xyVals = SplitPointsIntoXListAndYList(
          closestGrayPixelsInLine
        );
        List<double> xVals = xyVals.Item1;
        List<double> yVals = xyVals.Item2;

        double[] xValsArray = xVals.ToArray();
        double[] yValsArray = yVals.ToArray();

        double rSquared,
          yIntercept,
          slope;

        // Call the LinearRegression method
        LinearRegression(xValsArray, yValsArray, out rSquared, out yIntercept, out slope);

        // Create two points
        System.Drawing.PointF pt1 = new System.Drawing.PointF(0, (float)yIntercept);
        System.Drawing.PointF pt2 = new System.Drawing.PointF(1, (float)(slope + yIntercept));

        // Create a line as a tuple of two points
        Tuple<System.Drawing.PointF, System.Drawing.PointF> line = new Tuple<
          System.Drawing.PointF,
          System.Drawing.PointF
        >(pt1, pt2);

        line = EnsureCounterClockwise(line, centerX, centerY);

        var unitVector = GetOrthogonalVector(line);

        var averageX = closestGrayPixelsInPoint.Average(p => p.X);
        var averageY = closestGrayPixelsInPoint.Average(p => p.Y);
        var averagePoint = new System.Drawing.Point((int)averageX, (int)averageY);

        Vector2d vectorToCenter = new Vector2d(centerX - averagePoint.X, centerY - averagePoint.Y);

        var newPoint = new System.Windows.Point(
          point.X - vectorToCenter.X,
          point.Y - vectorToCenter.Y
        );

        var convertedBackPoint = ed.PointToWorld(newPoint, 0);

        MakeRecepBlockReference(doc, convertedBackPoint, unitVector);
      }
    }

    [CommandMethod("CheckViewport")]
    public void CheckViewport()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      using (Transaction trans = db.TransactionManager.StartTransaction())
      {
        try
        {
          // Check if we're in modelspace or paperspace
          if (db.TileMode)
          {
            // We're in modelspace
            ed.WriteMessage("\nCurrent Space: Modelspace");
          }
          else
          {
            // We're in paperspace
            ed.WriteMessage("\nCurrent Space: Paperspace");
          }

          trans.Commit();
        }
        catch (System.Exception ex)
        {
          ed.WriteMessage($"\nError: {ex.Message}");
          trans.Abort();
        }
      }
    }

    [CommandMethod("ImageAroundClick")]
    public void ImageAroundClick()
    {
      int WIDTH = 500;
      int HEIGHT = 500;
      int hue = 150;
      double saturation = 1.0;
      double value = 0.9608;

      Document doc = Application.DocumentManager.MdiActiveDocument;
      Editor ed = doc.Editor;

      PromptPointResult ppr = ed.GetPoint("\nClick a point: ");
      if (ppr.Status != PromptStatus.OK)
        return;
      Point3d clickPoint = ppr.Value;

      System.Windows.Point point = ed.PointToScreen(clickPoint, 0);
      System.Drawing.Point cursorPosition = System.Windows.Forms.Cursor.Position;

      int centerX = WIDTH / 2;
      int centerY = HEIGHT / 2;

      Rectangle captureRectangle = new Rectangle(
        cursorPosition.X - centerX,
        cursorPosition.Y - centerY,
        WIDTH,
        HEIGHT
      );

      using (Bitmap bitmap = new Bitmap(captureRectangle.Width, captureRectangle.Height))
      {
        using (Graphics g = Graphics.FromImage(bitmap))
        {
          g.CopyFromScreen(captureRectangle.Location, Point.Empty, captureRectangle.Size);
        }

        string imagePath = "capturedImage.png";
        bitmap.Save(imagePath, ImageFormat.Png);

        // Create the target HSV color
        Hsv targetHsv = new Hsv(hue, saturation * 255, value * 255);

        // Define the lower and upper bounds for the target color in HSV
        Hsv lowerBound = new Hsv(targetHsv.Hue - 10, 100, 100);
        Hsv upperBound = new Hsv(targetHsv.Hue + 10, 255, 255);

        // Convert the Hsv objects to ScalarArray
        ScalarArray lowerBoundScalar = new ScalarArray(
          new MCvScalar(lowerBound.Hue, lowerBound.Satuation, lowerBound.Value)
        );
        ScalarArray upperBoundScalar = new ScalarArray(
          new MCvScalar(upperBound.Hue, upperBound.Satuation, upperBound.Value)
        );

        // Load the captured image using Emgu CV
        Mat image = CvInvoke.Imread(imagePath);

        // Convert the image to the HSV color space
        Mat hsvImage = new Mat();
        CvInvoke.CvtColor(image, hsvImage, ColorConversion.Bgr2Hsv);

        // Create a binary mask based on the color range
        Mat mask = new Mat();
        CvInvoke.InRange(hsvImage, lowerBoundScalar, upperBoundScalar, mask);

        // Find contours in the binary mask
        VectorOfVectorOfPoint contours = new VectorOfVectorOfPoint();
        CvInvoke.FindContours(
          mask,
          contours,
          null,
          RetrType.External,
          ChainApproxMethod.ChainApproxSimple
        );

        // Iterate through the contours and draw boxes around them
        for (int i = 0; i < contours.Size; i++)
        {
          Rectangle rect = CvInvoke.BoundingRectangle(contours[i]);
          CvInvoke.Rectangle(image, rect, new Bgr(0, 255, 0).MCvScalar, 2);
        }

        // Save the image with boxes
        string outputImagePath = "outputImage.png";
        CvInvoke.Imwrite(outputImagePath, image);

        // Open the image with boxes
        System.Diagnostics.Process.Start(outputImagePath);
      }
    }

    [CommandMethod("Lighting")]
    public static void Lighting()
    {
      //var lightingForm = new INITIALIZE_LIGHTING_FORM();
      //lightingForm.Show();
      var lightingDialogWindow = new LightingDialogWindow();
      lightingDialogWindow.InitializeModal();
      lightingDialogWindow.Show();
    }

    [CommandMethod("DefineLightingSymbol")]
    public static void DefineLightingSymbol()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      GmepDatabase gmepDb = new GmepDatabase();

      PromptSelectionOptions promptSelectionOptions = new PromptSelectionOptions();
      promptSelectionOptions.MessageForAdding = "\nSelect Lighting:";
      PromptSelectionResult promptSelectionResult = ed.GetSelection(promptSelectionOptions);

      if (promptSelectionResult.Status == PromptStatus.OK)
      {
        SelectionSet selectionSet = promptSelectionResult.Value;
        PromptResult result = ed.GetString(
          "\nEnter scale (e.g., 1/4, 3/16) or press Enter to autoscale: "
        );
        string blockName = "";
        if (result.Status == PromptStatus.OK)
        {
          blockName = result.StringResult.Trim();
        }
        else
        {
          return;
        }
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

          if (bt.Has(blockName))
          {
            // Block exists
            MessageBox.Show(
              $"Block '{blockName}' already exists.",
              "Block Creation",
              MessageBoxButtons.OK,
              MessageBoxIcon.Information
            );
            tr.Commit();
            return;
          }

          TextStyleTable textStyleTable = (TextStyleTable)
            tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead);
          ObjectId gmepTextStyleId;
          if (textStyleTable.Has("gmep"))
          {
            gmepTextStyleId = textStyleTable["gmep"];
          }
          else
          {
            ed.WriteMessage("\nText style 'gmep' not found. Using default text style.");
            gmepTextStyleId = doc.Database.Textstyle;
          }

          BlockTableRecord newBlock = new BlockTableRecord();
          newBlock.Name = blockName;

          AttributeDefinition attrDefLightingCircuit = new AttributeDefinition();
          attrDefLightingCircuit.Position = new Point3d(-11.3082, 24.7579, 0);
          attrDefLightingCircuit.LockPositionInBlock = false;
          attrDefLightingCircuit.Tag = "LIGHTING_CIRCUIT";
          attrDefLightingCircuit.IsMTextAttributeDefinition = false;
          attrDefLightingCircuit.TextString = "#";
          attrDefLightingCircuit.Justify = AttachmentPoint.BaseLeft;
          attrDefLightingCircuit.Visible = true;
          attrDefLightingCircuit.Invisible = false;
          attrDefLightingCircuit.Constant = false;
          attrDefLightingCircuit.Height = 2;
          attrDefLightingCircuit.WidthFactor = 0.85;
          attrDefLightingCircuit.TextStyleId = gmepTextStyleId;
          attrDefLightingCircuit.Layer = "E-TEXT";

          newBlock.AppendEntity(attrDefLightingCircuit);

          AttributeDefinition attrDefLightingName = new AttributeDefinition();
          attrDefLightingName.Position = new Point3d(-12.6893, 27.3458, 0);
          attrDefLightingName.LockPositionInBlock = false;
          attrDefLightingName.Tag = "LIGHTING_NAME";
          attrDefLightingName.IsMTextAttributeDefinition = false;
          attrDefLightingName.TextString = "Name";
          attrDefLightingName.Justify = AttachmentPoint.BaseLeft;
          attrDefLightingName.Visible = true;
          attrDefLightingName.Invisible = false;
          attrDefLightingName.Constant = false;
          attrDefLightingName.Height = 2;
          attrDefLightingName.WidthFactor = 0.85;
          attrDefLightingName.TextStyleId = gmepTextStyleId;
          attrDefLightingName.Layer = "E-TEXT";

          newBlock.AppendEntity(attrDefLightingName);
          tr.AddNewlyCreatedDBObject(newBlock, true);

          bt.UpgradeOpen();

          bt.Add(newBlock);
          tr.AddNewlyCreatedDBObject(newBlock, true);

          ObjectId newBlockId = newBlock.ObjectId;

          ObjectIdCollection ids = new ObjectIdCollection();
          ObjectId[] idList = selectionSet.GetObjectIds();
          foreach (ObjectId id in idList)
          {
            ids.Add(id);
          }
          IdMapping mapping = new IdMapping();
          db.DeepCloneObjects(ids, newBlockId, mapping, false);

          tr.Commit();
        }
      }
    }

    [CommandMethod("PlaceControls")]
    public static void PlaceControls()
    {
      double scale = 12;

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
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Editor ed = doc.Editor;
      Database db = doc.Database;
      GmepDatabase gmepDb = new GmepDatabase();
      string projectId = gmepDb.GetProjectId(CADObjectCommands.GetProjectNoFromFileName());

      List<ElectricalEntity.LightingControl> lightingList = gmepDb.GetLightingControls(projectId);
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        LayerTable acLyrTbl;
        acLyrTbl = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

        string sLayerName = "E-LITE-CTRL";

        if (acLyrTbl.Has(sLayerName) == true)
        {
          db.Clayer = acLyrTbl[sLayerName];
          tr.Commit();
        }
      }
      int i = 0;
      foreach (ElectricalEntity.LightingControl control in lightingList)
      {
        ed.WriteMessage(
          "\nPlace "
            + (i + 1).ToString()
            + "/"
            + lightingList.Count.ToString()
            + " for '"
            + control.Name
            + "'"
        );
        string blockName = "GMEP LTG CTRL DIMMER";
        bool dimmerOccupancy = false;
        if (control.ControlType == "SWITCH")
        {
          blockName = "GMEP LTG CTRL SWITCH";
          if (control.HasOccupancy)
          {
            blockName = "GMEP LTG CTRL OCCUPANCY";
          }
        }
        else if (control.HasOccupancy)
        {
          dimmerOccupancy = true;
        }

        ObjectId blockId;
        Point3d point;
        double rotation = 0;
        try
        {
          using (Transaction tr = db.TransactionManager.StartTransaction())
          {
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

            BlockTableRecord block = (BlockTableRecord)
              tr.GetObject(bt[blockName], OpenMode.ForRead);
            BlockJig blockJig = new BlockJig();

            PromptResult res = blockJig.DragMe(block.ObjectId, out point);

            if (res.Status == PromptStatus.OK)
            {
              BlockTableRecord curSpace = (BlockTableRecord)
                tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

              BlockReference br = new BlockReference(point, block.ObjectId);

              RotateJig rotateJig = new RotateJig(br);
              PromptResult rotatePromptResult = ed.Drag(rotateJig);

              if (rotatePromptResult.Status != PromptStatus.OK)
              {
                return;
              }
              rotation = br.Rotation;
              br.ScaleFactors = new Scale3d(0.25 / scale);
              curSpace.AppendEntity(br);

              tr.AddNewlyCreatedDBObject(br, true);
              blockId = br.Id;

              foreach (ObjectId objId in block)
              {
                DBObject obj = tr.GetObject(objId, OpenMode.ForRead);
                AttributeDefinition attDef = obj as AttributeDefinition;
                if (attDef != null && !attDef.Constant)
                {
                  using (AttributeReference attRef = new AttributeReference())
                  {
                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                    if (attRef.Tag == "CONTROL_NAME")
                    {
                      attRef.Position = attDef.Position.TransformBy(br.BlockTransform);
                      attRef.TextString = control.Name;
                      attRef.Height = 0.0938 / scale * 6;
                      attRef.WidthFactor = 0.85;
                      attRef.HorizontalMode = TextHorizontalMode.TextLeft;
                      attRef.VerticalMode = TextVerticalMode.TextVerticalMid;
                      attRef.Justify = AttachmentPoint.BaseLeft;
                      attRef.Rotation = attRef.Rotation - rotation;
                    }
                    if (attRef.Tag == "D")
                    {
                      attRef.Position = attDef.Position.TransformBy(br.BlockTransform);
                      attRef.Height = 0.0938 / scale * 6;
                      attRef.WidthFactor = 0.85;
                      attRef.TextString = attRef.Tag;
                      attRef.HorizontalMode = TextHorizontalMode.TextLeft;
                      attRef.VerticalMode = TextVerticalMode.TextVerticalMid;
                      attRef.Justify = AttachmentPoint.BaseLeft;

                      Matrix3d rotationMatrix = Matrix3d.Rotation(
                        -rotation,
                        Vector3d.ZAxis,
                        br.Position
                      );
                      attRef.TransformBy(rotationMatrix);
                    }
                    br.AttributeCollection.AppendAttribute(attRef);
                    tr.AddNewlyCreatedDBObject(attRef, true);
                  }
                }
              }
            }
            else
            {
              return;
            }

            tr.Commit();
          }
          using (Transaction tr = db.TransactionManager.StartTransaction())
          {
            BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
            var modelSpace = (BlockTableRecord)
              tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
            BlockReference br = (BlockReference)tr.GetObject(blockId, OpenMode.ForWrite);
            DynamicBlockReferencePropertyCollection pc = br.DynamicBlockReferencePropertyCollection;
            foreach (DynamicBlockReferenceProperty prop in pc)
            {
              if (prop.PropertyName == "gmep_lighting_control_id" && prop.Value as string == "0")
              {
                prop.Value = control.Id;
              }
              if (prop.PropertyName == "gmep_lighting_control_tag" && prop.Value as string == "0")
              {
                prop.Value = control.Name;
              }
            }
            tr.Commit();
          }
        }
        catch (System.Exception ex)
        {
          ed.WriteMessage(ex.ToString());
        }
      }
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        LayerTable acLyrTbl;
        acLyrTbl = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

        string sLayerName = "E-CND1";

        if (acLyrTbl.Has(sLayerName) == true)
        {
          db.Clayer = acLyrTbl[sLayerName];
          tr.Commit();
        }
      }
    }

    [CommandMethod("ToggleEMLighting")]
    public static void ToggleEMLighting()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      GmepDatabase gmepDb = new GmepDatabase();

      PromptSelectionOptions pso2 = new PromptSelectionOptions();
      pso2.MessageForAdding = "\nSelect Lighting:";
      PromptSelectionResult psr2 = ed.GetSelection(pso2);

      if (psr2.Status == PromptStatus.OK)
      {
        SelectionSet ss = psr2.Value;
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          foreach (ObjectId id in ss.GetObjectIds())
          {
            string circuit = "";
            string control = "";
            string em = "";
            DBObject obj = tr.GetObject(id, OpenMode.ForWrite);
            BlockReference block = tr.GetObject(id, OpenMode.ForWrite) as BlockReference;
            foreach (
              DynamicBlockReferenceProperty property in block.DynamicBlockReferencePropertyCollection
            )
            {
              if (property.PropertyName == "gmep_lighting_control")
              {
                control = property.Value as string;
              }
              if (property.PropertyName == "gmep_lighting_circuit")
              {
                circuit = property.Value as string;
              }
              if (property.PropertyName == "Visibility1")
              {
                if (property.Value as string == "EM")
                {
                  property.Value = "Non-EM";
                }
                else
                {
                  property.Value = "EM";
                  em = ", EM";
                }
              }
            }
            foreach (ObjectId id2 in block.AttributeCollection)
            {
              AttributeReference attRef =
                tr.GetObject(id2, OpenMode.ForWrite) as AttributeReference;
              if (attRef.Tag == "LIGHTING_CIRCUIT")
              {
                attRef.TextString = circuit + control + em;
              }
            }
          }
          tr.Commit();
        }
      }
    }

    [CommandMethod("AssignLightingControl")]
    public static void AssignLightingControl()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      GmepDatabase gmepDb = new GmepDatabase();

      PromptSelectionOptions pso2 = new PromptSelectionOptions();
      pso2.MessageForAdding = "\nSelect a Lighting Control:";
      pso2.SingleOnly = true;
      PromptSelectionResult psr2 = ed.GetSelection(pso2);

      PromptSelectionOptions pso = new PromptSelectionOptions();
      pso.MessageForAdding = "\nSelect lighting fixtures: ";
      PromptSelectionResult psr = ed.GetSelection(pso);

      string control = "";

      if (psr2.Status == PromptStatus.OK)
      {
        SelectionSet ss = psr2.Value;
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          foreach (ObjectId id in ss.GetObjectIds())
          {
            DBObject obj = tr.GetObject(id, OpenMode.ForWrite);
            if (obj is BlockReference block)
            {
              foreach (
                DynamicBlockReferenceProperty property in block.DynamicBlockReferencePropertyCollection
              )
              {
                if (property.PropertyName == "gmep_lighting_control_tag")
                {
                  control = property.Value as string;
                }
              }
            }
          }
          tr.Commit();
        }
      }

      if (psr.Status == PromptStatus.OK)
      {
        SelectionSet ss = psr.Value;
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          foreach (ObjectId id in ss.GetObjectIds())
          {
            DBObject obj = tr.GetObject(id, OpenMode.ForWrite);
            if (obj is BlockReference block)
            {
              string circuit = "";
              string em = "";
              foreach (
                DynamicBlockReferenceProperty property in block.DynamicBlockReferencePropertyCollection
              )
              {
                if (property.PropertyName == "gmep_lighting_control")
                {
                  property.Value = control;
                }
                if (property.PropertyName == "gmep_lighting_circuit")
                {
                  circuit = property.Value as string;
                }
                if (property.PropertyName == "Visibility1")
                {
                  if (property.Value as string == "EM")
                  {
                    em = ", EM";
                  }
                }
              }
              foreach (ObjectId id2 in block.AttributeCollection)
              {
                AttributeReference attRef =
                  tr.GetObject(id2, OpenMode.ForWrite) as AttributeReference;
                if (attRef.Tag == "LIGHTING_CIRCUIT")
                {
                  attRef.TextString = circuit + control + em;
                }
              }
            }
          }
          tr.Commit();
        }
      }
    }

    [CommandMethod("AssignLightingCircuit")]
    public static void AssignLightingCircuit()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      GmepDatabase gmepDb = new GmepDatabase();

      string projectId = gmepDb.GetProjectId(CADObjectCommands.GetProjectNoFromFileName());
      List<Panel> panelList = gmepDb.GetPanels(projectId);
      List<ElectricalEntity.Equipment> equipmentList = gmepDb.GetEquipment(projectId);
      List<Transformer> transformerList = gmepDb.GetTransformers(projectId);
      Dictionary<string, List<string>> panelCircuits = new Dictionary<string, List<string>>();
      List<string> lightings = new List<string>();

      PromptKeywordOptions pko = new PromptKeywordOptions("");

      foreach (Panel panel in panelList)
      {
        if (panel.NumBreakers > 0)
        {
          pko.Keywords.Add(panel.Name + ":" + panel.Id);
          panelCircuits.Add(panel.Id, new List<string>());
          for (int i = 1; i <= panel.NumBreakers; i++)
          {
            panelCircuits[panel.Id].Add(i.ToString());
          }
        }
      }

      //Start removing Circuits from the dictionary, accounting for circuitnumber and pole.
      foreach (Panel panel in panelList)
      {
        if (panelCircuits.ContainsKey(panel.ParentId) && panel.Circuit != 0)
        {
          for (int i = 0; i < panel.Pole; i++)
          {
            panelCircuits[panel.ParentId].Remove((panel.Circuit + i * 2).ToString());
          }
        }
      }
      foreach (ElectricalEntity.Equipment equipment in equipmentList)
      {
        if (panelCircuits.ContainsKey(equipment.ParentId) && equipment.Circuit != 0)
        {
          for (int i = 0; i < equipment.Pole; i++)
          {
            panelCircuits[equipment.ParentId].Remove((equipment.Circuit + i * 2).ToString());
          }
        }
      }
      foreach (Transformer transformer in transformerList)
      {
        if (panelCircuits.ContainsKey(transformer.ParentId) && transformer.Circuit != 0)
        {
          for (int i = 0; i < transformer.Pole; i++)
          {
            panelCircuits[transformer.ParentId].Remove((transformer.Circuit + i * 2).ToString());
          }
        }
      }
      //end sorting circuits

      PromptSelectionResult psr = ed.GetSelection();

      //Prompt user for panel
      pko.Message = "\nAssign Panel:";
      PromptResult pr = ed.GetKeywords(pko);
      string result = pr.StringResult;
      var chosenPanel = result.Split(':')[1];
      var chosenPanelName = result.Split(':')[0];

      //Prompt user for circuit
      PromptKeywordOptions pko2 = new PromptKeywordOptions("");
      pko2.Message = "\nAssign Circuit:";
      foreach (string circuit in panelCircuits[chosenPanel])
      {
        pko2.Keywords.Add(circuit);
      }
      PromptResult pr2 = ed.GetKeywords(pko2);
      string result2 = pr2.StringResult;
      var chosenCircuit = int.Parse(result2);

      if (psr.Status == PromptStatus.OK)
      {
        SelectionSet ss = psr.Value;
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          foreach (ObjectId id in ss.GetObjectIds())
          {
            DBObject obj = tr.GetObject(id, OpenMode.ForWrite);
            if (obj is BlockReference block)
            {
              string circuit = "";
              string control = "";
              string em = "";
              foreach (
                DynamicBlockReferenceProperty property in block.DynamicBlockReferencePropertyCollection
              )
              {
                if (property.PropertyName == "gmep_lighting_fixture_id")
                {
                  var fixtureId = property.Value as string;
                  lightings.Add(fixtureId);
                }
              }
              foreach (
                DynamicBlockReferenceProperty property in block.DynamicBlockReferencePropertyCollection
              )
              {
                if (property.PropertyName == "gmep_lighting_parent_id")
                {
                  property.Value = chosenPanel;
                }
              }
              foreach (
                DynamicBlockReferenceProperty property in block.DynamicBlockReferencePropertyCollection
              )
              {
                if (property.PropertyName == "gmep_lighting_parent_name")
                {
                  property.Value = chosenPanelName;
                }
              }
              foreach (
                DynamicBlockReferenceProperty property in block.DynamicBlockReferencePropertyCollection
              )
              {
                if (property.PropertyName == "gmep_lighting_circuit")
                {
                  property.Value = chosenCircuit;
                  circuit = property.Value as string;
                }
              }
              foreach (
                DynamicBlockReferenceProperty property in block.DynamicBlockReferencePropertyCollection
              )
              {
                if (property.PropertyName == "gmep_lighting_control")
                {
                  control = property.Value as string;
                }
                if (property.PropertyName == "Visibility1")
                {
                  if (property.Value as string == "EM")
                  {
                    em = ", EM";
                  }
                }
              }
              foreach (ObjectId id3 in block.AttributeCollection)
              {
                AttributeReference attRef = (AttributeReference)
                  tr.GetObject(id3, OpenMode.ForWrite);
                if (attRef.Tag == "LIGHTING_CIRCUIT")
                {
                  attRef.TextString = circuit + control + em;
                }
              }
            }
          }

          gmepDb.InsertLightingEquipment(lightings, chosenPanel, chosenCircuit, projectId);
          tr.Commit();
        }
      }
    }

    [CommandMethod("DefineLightingLocation")]
    public static void DefineLightingLocation()
    {
      // Define the scale
      double scale = 12;
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

      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      GmepDatabase gmepDb = new GmepDatabase();
      string projectId = gmepDb.GetProjectId(CADObjectCommands.GetProjectNoFromFileName());
      List<LightingLocation> locationList = gmepDb.GetLightingLocations(projectId);
      List<LightingTimeClock> timeClockList = gmepDb.GetLightingTimeClocks(projectId);
      List<Panel> panelList = gmepDb.GetPanels(projectId);
      List<string> existingLocationIds = new List<string>();

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr =
          tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

        BlockTableRecord locationBlock = (BlockTableRecord)
          tr.GetObject(bt["LTG LOCATION"], OpenMode.ForRead);
        //Searching for existing locations
        foreach (ObjectId id in locationBlock.GetAnonymousBlockIds())
        {
          if (id.IsValid)
          {
            using (
              BlockTableRecord anonymousBtr = tr.GetObject(id, OpenMode.ForRead) as BlockTableRecord
            )
            {
              if (anonymousBtr != null)
              {
                foreach (ObjectId objId in anonymousBtr.GetBlockReferenceIds(true, false))
                {
                  var entity = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;
                  if (entity != null)
                  {
                    foreach (
                      DynamicBlockReferenceProperty prop in entity.DynamicBlockReferencePropertyCollection
                    )
                    {
                      if (prop.PropertyName == "lighting_location_id")
                      {
                        existingLocationIds.Add(prop.Value as string);
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }

      PromptKeywordOptions pko = new PromptKeywordOptions("");
      PromptKeywordOptions pko2 = new PromptKeywordOptions("");
      PromptKeywordOptions pko3 = new PromptKeywordOptions("");

      foreach (LightingLocation location in locationList)
      {
        if (!existingLocationIds.Contains(location.Id))
        {
          pko.Keywords.Add(location.LocationName.ToUpper() + ":" + location.Id);
        }
      }
      foreach (LightingTimeClock clock in timeClockList)
      {
        pko2.Keywords.Add(clock.Name.ToUpper() + ":" + clock.Id);
      }
      foreach (Panel panel in panelList)
      {
        pko3.Keywords.Add(panel.Name.ToUpper() + ":" + panel.Id);
      }

      PromptPointOptions ppo = new PromptPointOptions("\nSpecify start point: ");
      PromptPointResult ppr = ed.GetPoint(ppo);
      if (ppr.Status != PromptStatus.OK)
        return;

      Point3d startPoint = ppr.Value;
      PolyLineJig jig = new PolyLineJig(startPoint);

      // Loop to keep adding vertices until the polyline is closed
      while (true)
      {
        PromptResult res = ed.Drag(jig);
        if (res.Status == PromptStatus.OK)
        {
          jig.AddVertex(jig.CurrentPoint);
        }
        else if (res.Status == PromptStatus.Cancel)
        {
          break;
        }
      }

      Polyline polyline;
      using (Transaction trans = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr =
          trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
        polyline = jig.GetPolyline();

        if (polyline != null)
        {
          btr.AppendEntity(polyline);
          trans.AddNewlyCreatedDBObject(polyline, true);
          trans.Commit();
        }
      }

      pko.Message = "\nAssign Location:";
      PromptResult pr = ed.GetKeywords(pko);
      string result = pr.StringResult;
      var locationId = result.Split(':')[1];
      var locationName = result.Split(':')[0];

      Point3d point;
      ObjectId blockId;

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

        BlockTableRecord block = (BlockTableRecord)
          tr.GetObject(bt["LTG LOCATION"], OpenMode.ForRead);
        BlockJig blockJig = new BlockJig();

        PromptResult res = blockJig.DragMe(block.ObjectId, out point);

        if (res.Status == PromptStatus.OK)
        {
          BlockTableRecord curSpace = (BlockTableRecord)
            tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

          BlockReference br = new BlockReference(point, block.ObjectId);
          br.ScaleFactors = new Scale3d(0.25 / scale);

          curSpace.AppendEntity(br);
          tr.AddNewlyCreatedDBObject(br, true);
          blockId = br.Id;

          //Setting Attributes
          foreach (ObjectId objId in block)
          {
            DBObject obj = tr.GetObject(objId, OpenMode.ForRead);
            AttributeDefinition attDef = obj as AttributeDefinition;
            if (attDef != null && !attDef.Constant)
            {
              using (AttributeReference attRef = new AttributeReference())
              {
                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                attRef.Position = attDef.Position.TransformBy(br.BlockTransform);
                if (attDef.Tag == "NAME")
                {
                  attRef.TextString = locationName;
                }
                br.AttributeCollection.AppendAttribute(attRef);
                tr.AddNewlyCreatedDBObject(attRef, true);
              }
            }
          }
        }
        else
        {
          return;
        }

        tr.Commit();
      }
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
        var modelSpace = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
        BlockReference br = (BlockReference)tr.GetObject(blockId, OpenMode.ForWrite);
        DynamicBlockReferencePropertyCollection pc = br.DynamicBlockReferencePropertyCollection;

        LightingLocation matchingLocation = locationList.FirstOrDefault(x => x.Id == locationId);
        foreach (DynamicBlockReferenceProperty prop in pc)
        {
          if (prop.PropertyName == "lighting_location_id")
          {
            prop.Value = locationId;
          }
          if (prop.PropertyName == "outdoor")
          {
            if (matchingLocation.Outdoor)
              prop.Value = "True";
            else
              prop.Value = "False";
          }
          if (prop.PropertyName == "lighting_location_name")
          {
            prop.Value = locationName;
          }
        }
        tr.Commit();
      }

      //TimeClockPlacing
      pko2.Message = "\nWire To Time Clock:";
      PromptResult pr2 = ed.GetKeywords(pko2);
      string result2 = pr2.StringResult;
      var timeClockId = result2.Split(':')[1];
      var timeClockName = result2.Split(':')[0];

      Point3d point2;
      ObjectId blockId2 = ObjectId.Null;
      //bool timeClockPlaced = false;

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        BlockTableRecord btr =
          tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

        //Checking if timeclock is already placed
        foreach (ObjectId objId in btr)
        {
          Entity entity = tr.GetObject(objId, OpenMode.ForRead) as Entity;
          if (entity is BlockReference blockRef)
          {
            DynamicBlockReferencePropertyCollection pc =
              blockRef.DynamicBlockReferencePropertyCollection;
            foreach (DynamicBlockReferenceProperty prop in pc)
            {
              if (prop.PropertyName == "id")
              {
                if (prop.Value.ToString() == timeClockId)
                {
                  blockId2 = blockRef.ObjectId;
                  break;
                }
              }
            }
          }
        }

        BlockTableRecord block = (BlockTableRecord)
          tr.GetObject(bt["LTG TIMECLOCK"], OpenMode.ForRead);
        if (blockId2 == ObjectId.Null)
        {
          BlockJig blockJig = new BlockJig();
          PromptResult res = blockJig.DragMe(block.ObjectId, out point2);

          if (res.Status == PromptStatus.OK)
          {
            BlockTableRecord curSpace = (BlockTableRecord)
              tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

            BlockReference br = new BlockReference(point2, block.ObjectId);
            br.ScaleFactors = new Scale3d(0.25 / scale);

            curSpace.AppendEntity(br);
            tr.AddNewlyCreatedDBObject(br, true);
            blockId2 = br.Id;

            //Setting Attributes
            foreach (ObjectId objId in block)
            {
              DBObject obj = tr.GetObject(objId, OpenMode.ForRead);
              AttributeDefinition attDef = obj as AttributeDefinition;
              if (attDef != null && !attDef.Constant)
              {
                using (AttributeReference attRef = new AttributeReference())
                {
                  attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                  attRef.Position = attDef.Position.TransformBy(br.BlockTransform);
                  if (attDef.Tag == "NAME")
                  {
                    attRef.TextString = "TIMECLOCK " + timeClockName;
                  }
                  br.AttributeCollection.AppendAttribute(attRef);
                  tr.AddNewlyCreatedDBObject(attRef, true);
                }
              }
            }
          }
          else
          {
            return;
          }
        }

        tr.Commit();
      }
      pko3.Message = "\nPick Adjacent Panel For Time Clock:";
      PromptResult pr3 = ed.GetKeywords(pko3);
      string result3 = pr3.StringResult;
      var adjacentPanelId = result3.Split(':')[1];

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
        var modelSpace = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
        BlockReference br = (BlockReference)tr.GetObject(blockId2, OpenMode.ForWrite);
        DynamicBlockReferencePropertyCollection pc = br.DynamicBlockReferencePropertyCollection;

        LightingTimeClock matchingClock = timeClockList.FirstOrDefault(x => x.Id == timeClockId);
        foreach (DynamicBlockReferenceProperty prop in pc)
        {
          if (prop.PropertyName == "id")
          {
            prop.Value = timeClockId;
          }
          if (prop.PropertyName == "name")
          {
            prop.Value = timeClockName;
          }
          if (prop.PropertyName == "bypass_switch_name")
          {
            prop.Value = matchingClock.BypassSwitchName;
          }
          if (prop.PropertyName == "bypass_switch_location")
          {
            prop.Value = matchingClock.BypassSwitchLocation;
          }
          if (prop.PropertyName == "adjacent_panel_id")
          {
            prop.Value = adjacentPanelId;
          }
          if (prop.PropertyName == "voltage")
          {
            prop.Value = matchingClock.Voltage;
          }
        }
        tr.Commit();
      }

      //Assigning Timeclockid to location
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
        var modelSpace = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
        BlockReference br = (BlockReference)tr.GetObject(blockId, OpenMode.ForWrite);
        DynamicBlockReferencePropertyCollection pc = br.DynamicBlockReferencePropertyCollection;
        foreach (DynamicBlockReferenceProperty prop in pc)
        {
          if (prop.PropertyName == "timeclock_id")
          {
            prop.Value = timeClockId;
          }
        }
        tr.Commit();
      }
      SelectObjectsInsidePolyline(ed, db, polyline, locationId);
    }

    private static void SelectObjectsInsidePolyline(
      Editor ed,
      Database db,
      Polyline polyline,
      string locationId
    )
    {
      SelectionFilter filter = new SelectionFilter(
        new TypedValue[] { new TypedValue((int)DxfCode.Start, "INSERT") }
      );

      Extents3d extents = polyline.GeometricExtents;

      PromptSelectionResult psr = ed.SelectCrossingWindow(
        extents.MinPoint,
        extents.MaxPoint,
        filter
      );
      SelectionSet ss = psr.Value;
      if (psr.Status == PromptStatus.OK)
      {
        ed.WriteMessage($"\nNumber of objects selected: {ss.Count}");
      }
      else
      {
        ed.WriteMessage("\nNo objects selected.");
        return;
      }
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        foreach (ObjectId id in ss.GetObjectIds())
        {
          DBObject obj = tr.GetObject(id, OpenMode.ForWrite);
          if (obj is BlockReference block)
          {
            foreach (
              DynamicBlockReferenceProperty property in block.DynamicBlockReferencePropertyCollection
            )
            {
              if (property.PropertyName == "gmep_lighting_location_id")
              {
                property.Value = locationId;
              }
            }
          }
        }
        tr.Commit();
      }
    }

    [CommandMethod("LightingControlDiagram")]
    public static void LightingControlDiagram()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      List<LightingTimeClock> timeClocks = new List<LightingTimeClock>();
      List<LightingLocation> locations = new List<LightingLocation>();
      List<LightingFixture> lightings = new List<LightingFixture>();

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord timeClockBlock = (BlockTableRecord)
          tr.GetObject(bt["LTG TIMECLOCK"], OpenMode.ForRead);
        BlockTableRecord locationBlock = (BlockTableRecord)
          tr.GetObject(bt["LTG LOCATION"], OpenMode.ForRead);
        BlockTableRecord lightingBlock = (BlockTableRecord)
          tr.GetObject(bt["GMEP LTG 2X4"], OpenMode.ForRead);

        //Timeclocks
        foreach (ObjectId id in timeClockBlock.GetAnonymousBlockIds())
        {
          if (id.IsValid)
          {
            using (
              BlockTableRecord anonymousBtr = tr.GetObject(id, OpenMode.ForRead) as BlockTableRecord
            )
            {
              if (anonymousBtr != null)
              {
                foreach (ObjectId objId in anonymousBtr.GetBlockReferenceIds(true, false))
                {
                  var entity = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;

                  var timeClock = new LightingTimeClock("0", "", "", "", "", "");
                  var pc = entity.DynamicBlockReferencePropertyCollection;
                  foreach (DynamicBlockReferenceProperty prop in pc)
                  {
                    if (prop.PropertyName == "id")
                      timeClock.Id = prop.Value as string;
                    if (prop.PropertyName == "name")
                      timeClock.Name = prop.Value as string;
                    if (prop.PropertyName == "bypass_switch_name")
                      timeClock.BypassSwitchName = prop.Value as string;
                    if (prop.PropertyName == "bypass_switch_location")
                      timeClock.BypassSwitchLocation = prop.Value as string;
                    if (prop.PropertyName == "adjacent_panel_id")
                      timeClock.AdjacentPanelId = prop.Value as string;
                    if (prop.PropertyName == "voltage")
                      timeClock.Voltage = prop.Value as string;
                  }
                  if (timeClock.Id != "0")
                    timeClocks.Add(timeClock);
                }
              }
            }
          }
        }
        //Locations
        foreach (ObjectId id in locationBlock.GetAnonymousBlockIds())
        {
          if (id.IsValid)
          {
            using (
              BlockTableRecord anonymousBtr = tr.GetObject(id, OpenMode.ForRead) as BlockTableRecord
            )
            {
              if (anonymousBtr != null)
              {
                foreach (ObjectId objId in anonymousBtr.GetBlockReferenceIds(true, false))
                {
                  var entity = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;
                  var location = new LightingLocation("", "", false, "");
                  var pc = entity.DynamicBlockReferencePropertyCollection;
                  foreach (DynamicBlockReferenceProperty prop in pc)
                  {
                    if (prop.PropertyName == "lighting_location_id")
                      location.Id = prop.Value as string;
                    if (prop.PropertyName == "outdoor")
                      location.Outdoor = (prop.Value as string == "True");
                    if (prop.PropertyName == "lighting_location_name")
                      location.LocationName = prop.Value as string;
                    if (prop.PropertyName == "timeclock_id")
                      location.timeclock = prop.Value as string;
                  }
                  if (location.Id != "0")
                    locations.Add(location);
                }
              }
            }
          }
        }
        //Lighting
        foreach (ObjectId id in lightingBlock.GetAnonymousBlockIds())
        {
          if (id.IsValid)
          {
            using (
              BlockTableRecord anonymousBtr = tr.GetObject(id, OpenMode.ForRead) as BlockTableRecord
            )
            {
              if (anonymousBtr != null)
              {
                foreach (ObjectId objId in anonymousBtr.GetBlockReferenceIds(true, false))
                {
                  var entity = tr.GetObject(objId, OpenMode.ForRead) as BlockReference;
                  var lighting = new LightingFixture(
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    0,
                    0,
                    "",
                    1,
                    "",
                    "",
                    "",
                    "",
                    false,
                    0,
                    false,
                    false,
                    false,
                    0,
                    0,
                    0,
                    0,
                    0,
                    ""
                  );
                  var pc = entity.DynamicBlockReferencePropertyCollection;
                  foreach (DynamicBlockReferenceProperty prop in pc)
                  {
                    if (prop.PropertyName == "gmep_lighting_fixture_id")
                      lighting.Id = prop.Value as string;
                    if (prop.PropertyName == "gmep_lighting_parent_name")
                      lighting.ParentName = prop.Value as string;
                    if (prop.PropertyName == "gmep_lighting_location_id")
                      lighting.LocationId = prop.Value as string;
                    if (prop.PropertyName == "gmep_lighting_circuit")
                    {
                      if (int.TryParse(prop.Value.ToString(), out int circuit))
                      {
                        lighting.Circuit = circuit;
                      }
                    }
                    if (prop.PropertyName == "Visibility1")
                    {
                      //Note: Technically this should be a bool like IsEM/IsNotEm, but using EMCapable bool to determine if EM is toggled.
                      if (prop.Value as string == "EM")
                      {
                        lighting.EmCapable = true;
                      }
                      else
                      {
                        lighting.EmCapable = false;
                      }
                    }
                  }
                  foreach (ObjectId attId in entity.AttributeCollection)
                  {
                    var attRef = tr.GetObject(attId, OpenMode.ForWrite) as AttributeReference;
                    if (attRef != null && attRef.Tag == "LIGHTING_NAME")
                    {
                      lighting.Name = attRef.TextString;
                    }
                  }
                  if (lighting.Id != "0")
                    lightings.Add(lighting);
                }
              }
            }
          }
        }
        tr.Commit();
      }

      PromptKeywordOptions pko = new PromptKeywordOptions("");
      foreach (LightingTimeClock timeclock in timeClocks)
      {
        pko.Keywords.Add(timeclock.Name + ":" + timeclock.Id);
      }
      pko.Message = "\nSelect Time Clock For Diagram:";
      PromptResult pr = ed.GetKeywords(pko);
      string result = pr.StringResult;
      var timeClockId = result.Split(':')[1];

      LightingTimeClock chosenTimeClock = timeClocks.FirstOrDefault(x => x.Id == timeClockId);
      List<LightingLocation> newLocations = locations
        .Where(loc => loc.timeclock == chosenTimeClock.Id)
        .ToList();
      List<LightingFixture> newLightings = lightings
        .Where(lighting => newLocations.Any(loc => loc.Id == lighting.LocationId))
        .ToList();
      LightingControlDiagram diagram = new LightingControlDiagram(
        chosenTimeClock,
        newLocations,
        newLightings
      );
    }

    [CommandMethod("KMeans")]
    public void KMeans()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Editor ed = doc.Editor;
      Database db = doc.Database;
      var HEADER_HEIGHT = 32;

      ObjectId rectPolyID = new ObjectId();
      ObjectId hatchID = new ObjectId();

      string imagePath = "CapturedPolylineArea.png";

      PromptEntityResult per = PromptUserForPolyline(ed);
      if (per.Status != PromptStatus.OK)
        return;

      ClosePolyline(db, per);
      CreatePolylineAndHatchAroundInnerPolyline(ed, db, ref rectPolyID, ref hatchID, per);

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        Polyline poly = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Polyline;
        if (poly != null && poly.Closed)
        {
          Extents3d ext = poly.GeometricExtents;
          Point3d min = ext.MinPoint;
          Point3d max = ext.MaxPoint;

          // Convert the polyline's extents to screen coordinates
          var screenMin = ed.PointToScreen(min, 0);
          var screenMax = ed.PointToScreen(max, 0);

          // Get the top-left corner of the AutoCAD document window
          int screenHeight = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height;
          System.Windows.Point documentLocation = doc.Window.DeviceIndependentLocation;
          screenMin.Y += documentLocation.Y + HEADER_HEIGHT;
          screenMax.Y += documentLocation.Y + HEADER_HEIGHT;
          screenMin.X -= documentLocation.X;
          screenMax.X -= documentLocation.X;

          // Calculate the width and height of the captured area
          int width = (int)(screenMax.X - screenMin.X);
          int height = (int)(screenMin.Y - screenMax.Y);

          // Create a new bitmap to hold the captured image
          using (Bitmap bitmap = new Bitmap(width, height))
          {
            CaptureAScreenshot(
              ed,
              poly,
              screenMin,
              screenMax,
              documentLocation,
              width,
              height,
              bitmap
            );
            bitmap.Save(imagePath, ImageFormat.Png);
            var magentaObjects = LocateMagentaObjects(imagePath, ed, min, max, width, height);

            PutTextNearMagentaObjects(magentaObjects, ed, db);

            var numRooms = PromptForNumberOfRooms(ed);

            int numClusters = numRooms;
            double[][] data = magentaObjects
              .Select(obj => new double[] { obj.CenterNode.X, obj.CenterNode.Y })
              .ToArray();

            KMeans kmeans = new KMeans(numClusters);
            var clusters = kmeans.Learn(data);
            int[] labels = clusters.Decide(data);
            double[][] centroids = clusters.Centroids;

            var userPt = PromptUserForElectricalPanelPoint(ed);
            var orderedIndices = GetOrderedIndicesOfCentroidsFromUserClick(centroids, userPt);

            List<Edge> clusterEdges = new List<Edge>();
            HashSet<int> visitedClusters = new HashSet<int>();

            int startClusterIndex = orderedIndices[0];
            PrimVariation(
              startClusterIndex,
              magentaObjects,
              labels,
              visitedClusters,
              clusterEdges,
              userPt
            );

            foreach (var edge in clusterEdges)
            {
              CreateSplinesFromTriangulation(db, new List<Edge> { edge }, magentaObjects);
            }

            for (int i = 0; i < numClusters; i++)
            {
              var index = orderedIndices[i];
              List<MagentaObject> clusterObjects = magentaObjects
                .Where((obj, idx) => labels[idx] == index)
                .ToList();

              List<TriangleNet.Geometry.Vertex> vertices = clusterObjects
                .Select(obj => new TriangleNet.Geometry.Vertex(obj.CenterNode.X, obj.CenterNode.Y))
                .ToList();

              if (vertices.Count == 2)
              {
                var filteredEdges = new List<Edge>() { new Edge(vertices[0], vertices[1]) };
                CreateSplinesFromTriangulation(db, filteredEdges, clusterObjects);
              }
              else if (vertices.Count == 1) { }
              else
              {
                var triangulator = new Dwyer();

                var mesh = triangulator.Triangulate(vertices, new TriangleNet.Configuration());

                var edges = ConvertEdgesToPoints(mesh);

                var edgeStats = CalculateEdgeLengthStatistics(edges);
                double meanLength = edgeStats.mean;
                double stdDevLength = edgeStats.stdDev;

                List<Edge> filteredEdges = edges
                  .Where(edge =>
                  {
                    bool isLengthWithinRange = edge.Length() <= meanLength + stdDevLength;
                    bool isNode1Connected = CountConnectedEdges(edge.Point1, edges) >= 3;
                    bool isNode2Connected = CountConnectedEdges(edge.Point2, edges) >= 3;
                    return isLengthWithinRange || !isNode1Connected || !isNode2Connected;
                  })
                  .ToList();

                MagentaObject closestObject = FindClosestMagentaObject(userPt, clusterObjects);

                TriangleNet.Geometry.Vertex startNode = closestObject.CenterPointAsVertex();

                filteredEdges = BreadthFirstSearch(filteredEdges, startNode);

                CreateSplinesFromTriangulation(db, filteredEdges, magentaObjects);
              }
            }
          }
        }
        tr.Commit();
      }

      RemoveOuterPolylineAndHatch(db, rectPolyID, hatchID);
    }

    private void PutTextNearMagentaObjects(
      List<MagentaObject> magentaObjects,
      Editor ed,
      Database db
    )
    {
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
        LayerTableRecord ltr = tr.GetObject(lt["E-TXT1"], OpenMode.ForWrite) as LayerTableRecord;

        foreach (var obj in magentaObjects)
        {
          double maxX = obj.BoundaryPoints.Max(p => p.X);
          double maxY = obj.BoundaryPoints.Max(p => p.Y);
          Point3d position = new Point3d(maxX, maxY, 0);

          var dbTextId = Draw.Text("2a", 4.5, position, 0, false, "rpm");
          DBText dbText = tr.GetObject(dbTextId, OpenMode.ForWrite) as DBText;
          dbText.Layer = ltr.Name;
        }

        tr.Commit();
      }
    }

    private void PrimVariation(
      int startClusterIndex,
      List<MagentaObject> magentaObjects,
      int[] labels,
      HashSet<int> visitedClusters,
      List<Edge> clusterEdges,
      Point3d userPt
    )
    {
      visitedClusters.Add(startClusterIndex);

      while (visitedClusters.Count < labels.Distinct().Count())
      {
        double minDistance = double.MaxValue;
        int closestClusterIndex = -1;
        Edge closestEdge = null;

        foreach (int visitedClusterIndex in visitedClusters)
        {
          for (int i = 0; i < labels.Length; i++)
          {
            if (!visitedClusters.Contains(labels[i]))
            {
              Edge edge = FindClosestEdgeBetweenClusters(
                visitedClusterIndex,
                labels[i],
                magentaObjects,
                labels,
                userPt
              );
              if (edge != null && edge.Length() < minDistance)
              {
                minDistance = edge.Length();
                closestClusterIndex = labels[i];
                closestEdge = edge;
              }
            }
          }
        }

        if (closestClusterIndex != -1)
        {
          visitedClusters.Add(closestClusterIndex);
          clusterEdges.Add(closestEdge);
        }
      }
    }

    private Edge FindClosestEdgeBetweenClusters(
      int cluster1Label,
      int cluster2Label,
      List<MagentaObject> magentaObjects,
      int[] labels,
      Point3d userPt
    )
    {
      List<MagentaObject> cluster1Objects = magentaObjects
        .Where((obj, idx) => labels[idx] == cluster1Label)
        .ToList();
      List<MagentaObject> cluster2Objects = magentaObjects
        .Where((obj, idx) => labels[idx] == cluster2Label)
        .ToList();

      double minDistance = double.MaxValue;
      MagentaObject closestObject1 = null;
      MagentaObject closestObject2 = null;

      foreach (var obj1 in cluster1Objects)
      {
        foreach (var obj2 in cluster2Objects)
        {
          double distance1 = obj1.CenterNode.DistanceTo(userPt);
          double distance2 = obj2.CenterNode.DistanceTo(userPt);
          double distance = obj1.CenterNode.DistanceTo(obj2.CenterNode);
          double totalDistance = (distance * 0.7) + (distance1 * 0.15) + (distance2 * 0.15);

          if (totalDistance < minDistance)
          {
            minDistance = totalDistance;
            closestObject1 = obj1;
            closestObject2 = obj2;
          }
        }
      }

      if (closestObject1 != null && closestObject2 != null)
      {
        TriangleNet.Geometry.Vertex vertex1 = closestObject1.CenterPointAsVertex();
        TriangleNet.Geometry.Vertex vertex2 = closestObject2.CenterPointAsVertex();
        return new Edge(vertex1, vertex2);
      }

      return null;
    }

    private List<int> GetOrderedIndicesOfCentroidsFromUserClick(
      double[][] centroids,
      Point3d userPt
    )
    {
      List<int> indices = new List<int>();

      foreach (var centroid in centroids)
      {
        centroid[0] = Math.Abs(centroid[0] - userPt.X);
        centroid[1] = Math.Abs(centroid[1] - userPt.Y);
      }

      var orderedCentroids = centroids
        .Select((c, i) => new { Index = i, Distance = Math.Sqrt(c[0] * c[0] + c[1] * c[1]) })
        .OrderBy(c => c.Distance)
        .ToList();

      foreach (var centroid in orderedCentroids)
      {
        indices.Add(centroid.Index);
      }

      return indices;
    }

    private List<Edge> BreadthFirstSearch(List<Edge> edges, TriangleNet.Geometry.Vertex startNode)
    {
      var bfsEdges = new List<Edge>();
      var visited = new HashSet<TriangleNet.Geometry.Vertex>();
      var queue = new QuickGraph.Collections.Queue<TriangleNet.Geometry.Vertex>();

      visited.Add(startNode);
      queue.Enqueue(startNode);

      while (queue.Count > 0)
      {
        var currentNode = queue.Dequeue();

        foreach (
          var edge in edges.Where(e => e.Point1.Equals(currentNode) || e.Point2.Equals(currentNode))
        )
        {
          var neighborNode = edge.Point1.Equals(currentNode) ? edge.Point2 : edge.Point1;

          if (!visited.Contains(neighborNode))
          {
            visited.Add(neighborNode);
            queue.Enqueue(neighborNode);
            bfsEdges.Add(edge);
          }
        }
      }

      return bfsEdges;
    }

    private MagentaObject FindClosestMagentaObject(
      Point3d userPoint,
      List<MagentaObject> clusterObjects
    )
    {
      double minDistance = double.MaxValue;
      MagentaObject closestObject = null;

      foreach (var obj in clusterObjects)
      {
        double distance = userPoint.DistanceTo(obj.CenterNode);

        if (distance < minDistance)
        {
          minDistance = distance;
          closestObject = obj;
        }
      }

      return closestObject;
    }

    private static void CaptureAScreenshot(
      Editor ed,
      Polyline poly,
      System.Windows.Point screenMin,
      System.Windows.Point screenMax,
      System.Windows.Point documentLocation,
      int width,
      int height,
      Bitmap bitmap
    )
    {
      using (Graphics g = Graphics.FromImage(bitmap))
      {
        // Set the clipping boundary around the polyline
        GraphicsPath path = new GraphicsPath();
        // Get the polyline points
        int vertices = poly.NumberOfVertices;
        List<System.Drawing.PointF> polylinePoints = new List<System.Drawing.PointF>();
        for (int i = 0; i < vertices; i++)
        {
          Point3d point = poly.GetPoint3dAt(i);
          System.Windows.Point screenPoint = ed.PointToScreen(point, 0);
          screenPoint.X -= documentLocation.X;
          screenPoint.Y -= documentLocation.Y;
          polylinePoints.Add(
            new System.Drawing.PointF(
              (float)(screenPoint.X - screenMin.X),
              (float)(screenMin.Y - screenPoint.Y)
            )
          );
        }
        path.AddPolygon(polylinePoints.ToArray());
        g.SetClip(path);

        // Capture the contents of the AutoCAD window inside the polyline
        g.CopyFromScreen(
          (int)screenMin.X,
          (int)screenMax.Y,
          0,
          0,
          new System.Drawing.Size(width, height)
        );
      }
    }

    private static void RemoveOuterPolylineAndHatch(
      Database db,
      ObjectId rectPolyID,
      ObjectId hatchID
    )
    {
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        if (rectPolyID != null)
        {
          DBObject outerPoly = tr.GetObject(rectPolyID, OpenMode.ForWrite);
          outerPoly.Erase();
        }

        if (hatchID != null)
        {
          DBObject hatch = tr.GetObject(hatchID, OpenMode.ForWrite);
          hatch.Erase();
        }

        tr.Commit();
      }
    }

    private static void CreatePolylineAndHatchAroundInnerPolyline(
      Editor ed,
      Database db,
      ref ObjectId rectPolyID,
      ref ObjectId hatchID,
      PromptEntityResult per
    )
    {
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        // Get the selected polyline
        Polyline poly = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Polyline;

        if (poly != null && poly.Closed)
        {
          // Get the polyline's bounding box
          Extents3d ext = poly.GeometricExtents;
          ed.Zoom(ext);
          Point3d min = ext.MinPoint;
          Point3d max = ext.MaxPoint;

          rectPolyID = CreateOuterPolyline(db, tr, ref min, ref max);

          ObjectId[] polys = new ObjectId[] { poly.ObjectId, rectPolyID };

          hatchID = Hatch(polys);

          ed.UpdateScreen();
        }
        tr.Commit();
      }
    }

    private static PromptEntityResult PromptUserForPolyline(Editor ed)
    {
      PromptEntityOptions peo = new PromptEntityOptions("");
      peo.SetRejectMessage("\nSelected object is not a closed polyline.");
      peo.AddAllowedClass(typeof(Polyline), true);
      PromptEntityResult per = ed.GetEntity(peo);
      return per;
    }

    private static Point3d PromptUserForElectricalPanelPoint(Editor ed)
    {
      PromptPointOptions ppo = new PromptPointOptions("\nClick on the electrical panel: ");
      ppo.AllowNone = false; // User must select a point

      PromptPointResult ppr = ed.GetPoint(ppo);
      if (ppr.Status != PromptStatus.OK)
      {
        ed.WriteMessage("\nPrompt was cancelled.");
        return Point3d.Origin; // Return the origin point if the prompt was cancelled
      }

      return ppr.Value;
    }

    private static void ClosePolyline(Database db, PromptEntityResult per)
    {
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        // Get the selected polyline
        Polyline poly = tr.GetObject(per.ObjectId, OpenMode.ForWrite) as Polyline;
        if (poly != null)
        {
          if (!poly.Closed)
          {
            poly.Closed = true;
            tr.Commit();
          }
        }
      }
    }

    private static int PromptForNumberOfRooms(Editor ed)
    {
      int numRooms = 0;
      bool validInput = false;

      while (!validInput)
      {
        PromptIntegerOptions pio = new PromptIntegerOptions("\nEnter the number of rooms: ");
        pio.AllowNegative = false;
        pio.AllowZero = false;
        pio.DefaultValue = 1;
        pio.LowerLimit = 1;
        pio.UpperLimit = 100;

        PromptIntegerResult pir = ed.GetInteger(pio);

        if (pir.Status == PromptStatus.OK)
        {
          numRooms = pir.Value;
          validInput = true;
        }
        else if (pir.Status == PromptStatus.Cancel)
        {
          throw new OperationCanceledException("User canceled the operation.");
        }
        else
        {
          ed.WriteMessage("\nInvalid input. Please enter a valid number of rooms.");
        }
      }

      return numRooms;
    }

    private int CountConnectedEdges(TriangleNet.Geometry.Vertex vertex, List<Edge> edges)
    {
      int count = 0;
      foreach (var edge in edges)
      {
        if (edge.Point1.Equals(vertex) || edge.Point2.Equals(vertex))
        {
          count++;
        }
      }
      return count;
    }

    private (double mean, double stdDev) CalculateEdgeLengthStatistics(List<Edge> edges)
    {
      double sum = 0;
      foreach (var edge in edges)
      {
        sum += edge.Length();
      }
      double mean = sum / edges.Count;

      double sumSquaredDiff = 0;
      foreach (var edge in edges)
      {
        double diff = edge.Length() - mean;
        sumSquaredDiff += diff * diff;
      }
      double variance = sumSquaredDiff / edges.Count;
      double stdDev = Math.Sqrt(variance);

      return (mean, stdDev);
    }

    private List<Edge> ConvertEdgesToPoints(TriangleNet.Meshing.IMesh mesh)
    {
      List<Edge> edges = new List<Edge>();

      foreach (var edge in mesh.Edges)
      {
        var vert1 = mesh.Vertices.First(vertex => vertex.ID == edge.P0);
        var vert2 = mesh.Vertices.First(vertex => vertex.ID == edge.P1);

        edges.Add(new Edge(vert1, vert2));
      }

      return edges;
    }

    private void CreateSplinesFromTriangulation(
      Database db,
      List<Edge> edges,
      List<MagentaObject> magentaObjects
    )
    {
      using (Transaction trSplines = db.TransactionManager.StartTransaction())
      {
        BlockTable blockTable =
          trSplines.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord modelSpace =
          trSplines.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite)
          as BlockTableRecord;

        foreach (var edge in edges)
        {
          TriangleNet.Geometry.Vertex vertex1 = edge.Point1;
          TriangleNet.Geometry.Vertex vertex2 = edge.Point2;

          MagentaObject magentaObject1 = magentaObjects.FirstOrDefault(obj =>
            obj.CenterNode.X == vertex1.X && obj.CenterNode.Y == vertex1.Y
          );
          MagentaObject magentaObject2 = magentaObjects.FirstOrDefault(obj =>
            obj.CenterNode.X == vertex2.X && obj.CenterNode.Y == vertex2.Y
          );

          if (magentaObject1 != null && magentaObject2 != null)
          {
            Point3d closestPoint1 = FindClosestPoint(
              magentaObject1.BoundaryPoints,
              magentaObject2.CenterNode
            );
            Point3d closestPoint2 = FindClosestPoint(
              magentaObject2.BoundaryPoints,
              magentaObject1.CenterNode
            );

            Spline spline = CreateSpline(closestPoint1, closestPoint2);

            modelSpace.AppendEntity(spline);
            trSplines.AddNewlyCreatedDBObject(spline, true);
          }
        }

        trSplines.Commit();
      }
    }

    private static Spline CreateSpline(Point3d startPoint, Point3d endPoint)
    {
      Spline spline = new Spline(
        new Point3dCollection(new[] { startPoint, endPoint }),
        Vector3d.ZAxis,
        Vector3d.ZAxis,
        3,
        0.0
      );

      // Calculate the midpoint between the start and end points
      Point3d midPoint = new Point3d(
        (startPoint.X + endPoint.X) / 2,
        (startPoint.Y + endPoint.Y) / 2,
        0
      );

      // Calculate the direction vector from start point to end point
      Vector3d direction = endPoint - startPoint;

      // Calculate the perpendicular vector to the direction vector
      Vector3d perpendicular = new Vector3d(-direction.Y, direction.X, 0);

      // Normalize the perpendicular vector
      perpendicular = perpendicular.GetNormal();

      // Calculate the distance between the start and end points
      double distance = startPoint.DistanceTo(endPoint);

      // Calculate the offset distance for the control points
      double offsetDistance = distance * 0.2; // Adjust this value to control the curvature

      // Calculate the control points by adding the scaled perpendicular vector to the midpoint
      Point3d controlPoint1 = midPoint + perpendicular * offsetDistance;

      // Set the second control point to be the same as the end point
      Point3d controlPoint2 = endPoint;

      // Add the control points to the spline
      spline.SetControlPointAt(1, controlPoint1);
      spline.SetControlPointAt(2, controlPoint2);

      return spline;
    }

    private static Point3d FindClosestPoint(List<Point3d> boundaryPoints, Point3d targetNode)
    {
      Point3d closestPoint = Point3d.Origin;
      double minDistance = double.MaxValue;

      foreach (var point in boundaryPoints)
      {
        double distance = point.DistanceTo(new Point3d(targetNode.X, targetNode.Y, 0));
        if (distance < minDistance)
        {
          minDistance = distance;
          closestPoint = point;
        }
      }

      return closestPoint;
    }

    public Tuple<System.Drawing.PointF, System.Drawing.PointF> EnsureCounterClockwise(
      Tuple<System.Drawing.PointF, System.Drawing.PointF> line,
      float centerX,
      float centerY
    )
    {
      var pt1 = line.Item1;
      var pt2 = line.Item2;

      // Calculate the angles of the points from the positive x-axis
      var angle1 = Math.Atan2(pt1.Y - centerY, pt1.X - centerX);
      var angle2 = Math.Atan2(pt2.Y - centerY, pt2.X - centerX);

      // If the angle of pt2 is less than the angle of pt1, swap the points
      if (angle2 < angle1)
      {
        var temp = pt1;
        pt1 = pt2;
        pt2 = temp;
      }

      return new Tuple<System.Drawing.PointF, System.Drawing.PointF>(pt1, pt2);
    }

    private Tuple<List<double>, List<double>> SplitPointsIntoXListAndYList(
      List<System.Drawing.Point> points
    )
    {
      List<double> xVals = new List<double>();
      List<double> yVals = new List<double>();

      foreach (var point in points)
      {
        xVals.Add(point.X);
        yVals.Add(point.Y);
      }

      return new Tuple<List<double>, List<double>>(xVals, yVals);
    }

    private Vector3d GetOrthogonalVector(Tuple<System.Drawing.PointF, System.Drawing.PointF> line)
    {
      // Calculate the difference in x and y coordinates between the end point and the start point
      double dx = line.Item2.X - line.Item1.X;
      double dy = line.Item2.Y - line.Item1.Y;

      // Create a vector from the differences
      Vector3d vector = new Vector3d(dx, dy, 0);

      // Create a vector pointing in the 0,0,1 direction
      Vector3d upVector = new Vector3d(0, 0, 1);

      // Calculate the cross product of the two vectors
      Vector3d crossProduct = vector.CrossProduct(upVector);

      // Normalize the cross product to get a unit vector
      Vector3d unitVector = crossProduct.GetNormal();

      return unitVector;
    }

    public static void LinearRegression(
      double[] xVals,
      double[] yVals,
      out double rSquared,
      out double yIntercept,
      out double slope
    )
    {
      if (xVals.Length != yVals.Length)
      {
        throw new System.Exception("Input values should be with the same length.");
      }

      double sumOfX = 0;
      double sumOfY = 0;
      double sumOfXSq = 0;
      double sumOfYSq = 0;
      double sumCodeviates = 0;

      for (var i = 0; i < xVals.Length; i++)
      {
        var x = xVals[i];
        var y = yVals[i];
        sumCodeviates += x * y;
        sumOfX += x;
        sumOfY += y;
        sumOfXSq += x * x;
        sumOfYSq += y * y;
      }

      var count = xVals.Length;
      var ssX = sumOfXSq - ((sumOfX * sumOfX) / count);
      var ssY = sumOfYSq - ((sumOfY * sumOfY) / count);

      var rNumerator = (count * sumCodeviates) - (sumOfX * sumOfY);
      var rDenom = (count * sumOfXSq - (sumOfX * sumOfX)) * (count * sumOfYSq - (sumOfY * sumOfY));
      var sCo = sumCodeviates - ((sumOfX * sumOfY) / count);

      var meanX = sumOfX / count;
      var meanY = sumOfY / count;
      var dblR = rNumerator / Math.Sqrt(rDenom);

      rSquared = dblR * dblR;
      yIntercept = meanY - ((sCo / ssX) * meanX);
      slope = sCo / ssX;
    }

    private List<System.Drawing.Point> FindClosestGrayPixelsWithBlackBorders(
      List<System.Drawing.Point> grayPixels,
      Bitmap bitmap,
      int amount,
      int centerImageX,
      int centerImageY
    )
    {
      var grayPixelsWithBlackBorders = FindGrayPixelsWithBlackBorders(
        grayPixels,
        bitmap,
        centerImageX,
        centerImageY
      );

      var distances = grayPixelsWithBlackBorders.Select(p => new
      {
        Pixel = p,
        Distance = Math.Sqrt(Math.Pow(p.X - centerImageX, 2) + Math.Pow(p.Y - centerImageY, 2)),
      });

      var closestGrayPixels = distances
        .OrderBy(d => d.Distance)
        .Take(amount)
        .Select(d => d.Pixel)
        .ToList();

      return closestGrayPixels;
    }

    private List<System.Drawing.Point> FindGrayPixelsWithBlackBorders(
      List<System.Drawing.Point> grayPixels,
      Bitmap bitmap,
      int centerImageX,
      int centerImageY
    )
    {
      List<System.Drawing.Point> grayPixelsWithBlackBorders = new List<System.Drawing.Point>();

      foreach (var grayPixel in grayPixels)
      {
        List<System.Drawing.Point> blackBorderPixels = new List<System.Drawing.Point>();

        // Check the pixel to the left
        if (grayPixel.X > 0 && bitmap.GetPixel(grayPixel.X - 1, grayPixel.Y).R == 0)
        {
          blackBorderPixels.Add(new System.Drawing.Point(grayPixel.X - 1, grayPixel.Y));
        }
        // Check the pixel to the right
        if (grayPixel.X < bitmap.Width - 1 && bitmap.GetPixel(grayPixel.X + 1, grayPixel.Y).R == 0)
        {
          blackBorderPixels.Add(new System.Drawing.Point(grayPixel.X + 1, grayPixel.Y));
        }
        // Check the pixel above
        if (grayPixel.Y > 0 && bitmap.GetPixel(grayPixel.X, grayPixel.Y - 1).R == 0)
        {
          blackBorderPixels.Add(new System.Drawing.Point(grayPixel.X, grayPixel.Y - 1));
        }
        // Check the pixel below
        if (grayPixel.Y < bitmap.Height - 1 && bitmap.GetPixel(grayPixel.X, grayPixel.Y + 1).R == 0)
        {
          blackBorderPixels.Add(new System.Drawing.Point(grayPixel.X, grayPixel.Y + 1));
        }

        double grayPixelDistance = Math.Sqrt(
          Math.Pow(grayPixel.X - centerImageX, 2) + Math.Pow(grayPixel.Y - centerImageY, 2)
        );

        foreach (var blackPixel in blackBorderPixels)
        {
          double blackPixelDistance = Math.Sqrt(
            Math.Pow(blackPixel.X - centerImageX, 2) + Math.Pow(blackPixel.Y - centerImageY, 2)
          );

          if (blackPixelDistance < grayPixelDistance)
          {
            grayPixelsWithBlackBorders.Add(grayPixel);
            break;
          }
        }
      }
      return grayPixelsWithBlackBorders;
    }

    private void MakeRecepBlockReference(
      Document doc,
      Point3d convertedBackPoint,
      Vector3d unitVector
    )
    {
      var blockName = "RECEP";
      using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
      {
        BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

        BlockReference br = new BlockReference(convertedBackPoint, bt[blockName]);

        // Calculate the angle of rotation from the vector
        double angle = Math.Atan2(unitVector.Y, unitVector.X);

        // Correct the angle for AutoCAD's coordinate system
        angle = Math.PI / 2 - angle;

        if (angle < 0)
        {
          angle += 2 * Math.PI;
        }

        // Set the rotation of the block reference
        br.Rotation = angle;

        btr.AppendEntity(br);
        tr.AddNewlyCreatedDBObject(br, true);

        tr.Commit();
      }
    }

    private List<System.Drawing.Point> FindGrayPixelsAroundClosest(
      List<System.Drawing.Point> grayPixels,
      Bitmap bitmap,
      int centerImageX,
      int centerImageY,
      System.Drawing.Point? closestGrayPixel
    )
    {
      if (!closestGrayPixel.HasValue)
      {
        return new List<System.Drawing.Point>();
      }

      var surroundingGrayPixels = new List<System.Drawing.Point>();

      for (int x = closestGrayPixel.Value.X - 2; x <= closestGrayPixel.Value.X + 2; x++)
      {
        for (int y = closestGrayPixel.Value.Y - 2; y <= closestGrayPixel.Value.Y + 2; y++)
        {
          if (x >= 0 && x < bitmap.Width && y >= 0 && y < bitmap.Height)
          {
            Color pixelColor = bitmap.GetPixel(x, y);
            if (pixelColor.R == pixelColor.G && pixelColor.G == pixelColor.B && pixelColor.R != 0)
            {
              surroundingGrayPixels.Add(new System.Drawing.Point(x, y));
            }
          }
        }
      }

      return surroundingGrayPixels;
    }

    private System.Drawing.Point? FindClosestGrayPixelWithBlackBorder(
      List<System.Drawing.Point> grayPixels,
      Bitmap bitmap,
      int centerImageX,
      int centerImageY
    )
    {
      System.Drawing.Point? closestGrayPixel = null;
      double closestDistance = double.MaxValue;

      foreach (var grayPixel in grayPixels)
      {
        // Check the surrounding pixels for a black one
        var surroundingPoints = new List<System.Drawing.Point>
        {
          new System.Drawing.Point(grayPixel.X - 1, grayPixel.Y),
          new System.Drawing.Point(grayPixel.X + 1, grayPixel.Y),
          new System.Drawing.Point(grayPixel.X, grayPixel.Y - 1),
          new System.Drawing.Point(grayPixel.X, grayPixel.Y + 1),
        };

        foreach (var point in surroundingPoints)
        {
          if (point.X >= 0 && point.X < bitmap.Width && point.Y >= 0 && point.Y < bitmap.Height)
          {
            Color pixelColor = bitmap.GetPixel(point.X, point.Y);
            if (pixelColor.R == 0 && pixelColor.G == 0 && pixelColor.B == 0) // if the pixel is black
            {
              double distance = Math.Sqrt(
                Math.Pow(grayPixel.X - centerImageX, 2) + Math.Pow(grayPixel.Y - centerImageY, 2)
              );
              if (distance < closestDistance)
              {
                closestDistance = distance;
                closestGrayPixel = grayPixel;
              }
              break;
            }
          }
        }
      }

      return closestGrayPixel;
    }

    private static List<MagentaObject> LocateMagentaObjects(
      string imagePath,
      Editor editor,
      Point3d min,
      Point3d max,
      int width,
      int height
    )
    {
      int hue = 150;
      double saturation = 1.0;
      double value = 0.9608;

      // Create the target HSV color
      Hsv targetHsv = new Hsv(hue, saturation * 255, value * 255);

      // Define the lower and upper bounds for the target color in HSV
      Hsv lowerBound = new Hsv(targetHsv.Hue - 10, 100, 100);
      Hsv upperBound = new Hsv(targetHsv.Hue + 10, 255, 255);

      // Convert the Hsv objects to ScalarArray
      ScalarArray lowerBoundScalar = new ScalarArray(
        new MCvScalar(lowerBound.Hue, lowerBound.Satuation, lowerBound.Value)
      );
      ScalarArray upperBoundScalar = new ScalarArray(
        new MCvScalar(upperBound.Hue, upperBound.Satuation, upperBound.Value)
      );

      // Load the captured image using Emgu CV
      Mat image = CvInvoke.Imread(imagePath);

      // Convert the image to the HSV color space
      Mat hsvImage = new Mat();
      CvInvoke.CvtColor(image, hsvImage, ColorConversion.Bgr2Hsv);

      // Create a binary mask based on the color range
      Mat mask = new Mat();
      CvInvoke.InRange(hsvImage, lowerBoundScalar, upperBoundScalar, mask);

      // Find contours in the binary mask
      VectorOfVectorOfPoint contours = new VectorOfVectorOfPoint();
      CvInvoke.FindContours(
        mask,
        contours,
        null,
        RetrType.External,
        ChainApproxMethod.ChainApproxSimple
      );

      List<MagentaObject> magentaObjects = new List<MagentaObject>();

      // Iterate through the contours and create MagentaObject instances
      for (int i = 0; i < contours.Size; i++)
      {
        double epsilon = 0.02 * CvInvoke.ArcLength(contours[i], true);
        VectorOfPoint approx = new VectorOfPoint();
        CvInvoke.ApproxPolyDP(contours[i], approx, epsilon, true);

        List<Point3d> boundaryPointsAutoCAD = new List<Point3d>();

        for (int j = 0; j < approx.Size; j++)
        {
          Point pixelPoint = new Point(approx[j].X, approx[j].Y);
          Point3d autoCADPoint = ConvertPixelToAutoCAD(pixelPoint, editor, min, max, width, height);
          boundaryPointsAutoCAD.Add(autoCADPoint);
        }

        UpdateBoundaryPoints(boundaryPointsAutoCAD, min, max);
        MagentaObject magentaObject = new MagentaObject(boundaryPointsAutoCAD);
        magentaObjects.Add(magentaObject);

        CvInvoke.DrawContours(
          image,
          new VectorOfVectorOfPoint(approx),
          -1,
          new Bgr(0, 255, 0).MCvScalar,
          2
        );
      }

      return magentaObjects;
    }

    private static void UpdateBoundaryPoints(
      List<Point3d> boundaryPointsAutoCAD,
      Point3d min,
      Point3d max
    )
    {
      double baseScaleFactor = 1000; // Adjust this value based on your desired base scale factor

      // Calculate the extents of the model space
      double extentX = max.X - min.X;
      double extentY = max.Y - min.Y;

      // Calculate the scale factor based on the extents
      double scaleFactor = Math.Max(extentX, extentY) / baseScaleFactor;

      // Calculate the adjusted deltaX and deltaY values proportional to the scale factor
      double deltaX = 1 * scaleFactor;
      double deltaY = 0.2466 * scaleFactor / 1000;

      for (int i = 0; i < boundaryPointsAutoCAD.Count; i++)
      {
        Point3d point = boundaryPointsAutoCAD[i];
        point = new Point3d(point.X + deltaX, point.Y + deltaY, 0);
        boundaryPointsAutoCAD[i] = point;
      }
    }

    private static Point3d ConvertPixelToAutoCAD(
      Point pixelPoint,
      Editor editor,
      Point3d min,
      Point3d max,
      int width,
      int height
    )
    {
      double worldX = min.X + (pixelPoint.X * (max.X - min.X)) / width;
      double worldY = max.Y - (pixelPoint.Y * (max.Y - min.Y)) / height;

      return new Point3d(worldX, worldY, 0);
    }

    public static ObjectId Hatch(
      ObjectId[] loopIds,
      string hatchName = "SOLID",
      double scale = 1,
      double angle = 0,
      bool associative = false
    )
    {
      var db = GetDatabase(loopIds);
      using (var trans = db.TransactionManager.StartTransaction())
      {
        var hatch = new Hatch();
        var space = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
        ObjectId hatchId = space.AppendEntity(hatch);
        trans.AddNewlyCreatedDBObject(hatch, true);

        hatch.SetDatabaseDefaults();
        hatch.Normal = new Vector3d(0, 0, 1);
        hatch.Elevation = 0.0;
        hatch.Associative = associative;
        hatch.PatternScale = scale;
        hatch.SetHatchPattern(HatchPatternType.PreDefined, hatchName);
        hatch.ColorIndex = 0;
        hatch.PatternAngle = angle;
        hatch.HatchStyle = Autodesk.AutoCAD.DatabaseServices.HatchStyle.Outer;
        for (int i = 0; i < loopIds.Length; i++)
        {
          ObjectId loop = loopIds[i];
          hatch.AppendLoop(HatchLoopTypes.External, new ObjectIdCollection(new[] { loop }));
        }
        hatch.EvaluateHatch(true);

        trans.Commit();
        return hatchId;
      }
    }

    internal static Database GetDatabase(IEnumerable<ObjectId> objectIds)
    {
      return objectIds.Select(id => id.Database).Distinct().Single();
    }

    private static ObjectId CreateOuterPolyline(
      Database db,
      Transaction tr,
      ref Point3d min,
      ref Point3d max
    )
    {
      // Create a rectangular polyline around the selected polyline
      double offset = 1.0; // Adjust the offset value as needed
      Point2d pt1 = new Point2d(min.X - offset, min.Y - offset);
      Point2d pt2 = new Point2d(max.X + offset, min.Y - offset);
      Point2d pt3 = new Point2d(max.X + offset, max.Y + offset);
      Point2d pt4 = new Point2d(min.X - offset, max.Y + offset);

      Polyline rectPoly = new Polyline();
      rectPoly.AddVertexAt(0, pt1, 0, 0, 0);
      rectPoly.AddVertexAt(1, pt2, 0, 0, 0);
      rectPoly.AddVertexAt(2, pt3, 0, 0, 0);
      rectPoly.AddVertexAt(3, pt4, 0, 0, 0);
      rectPoly.Closed = true;

      // Add the rectangular polyline to the database
      BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
      ObjectId polylineId = btr.AppendEntity(rectPoly);
      tr.AddNewlyCreatedDBObject(rectPoly, true);

      return polylineId;
    }

    public static void SaveDataToJsonFileOnDesktop(
      object data,
      string fileName,
      bool noOverride = false
    )
    {
      string jsonData = JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented);
      string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
      string fullPath = Path.Combine(desktopPath, fileName);

      if (noOverride && File.Exists(fullPath))
      {
        int fileNumber = 1;
        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
        string fileExtension = Path.GetExtension(fileName);

        while (File.Exists(fullPath))
        {
          string newFileName = $"{fileNameWithoutExtension} ({fileNumber}){fileExtension}";
          fullPath = Path.Combine(desktopPath, newFileName);
          fileNumber++;
        }
      }

      File.WriteAllText(fullPath, jsonData);
    }
  }

  public class Edge
  {
    public TriangleNet.Geometry.Vertex Point1 { get; set; }
    public TriangleNet.Geometry.Vertex Point2 { get; set; }

    public Edge(TriangleNet.Geometry.Vertex vertex1, TriangleNet.Geometry.Vertex vertex2)
    {
      Point1 = vertex1;
      Point2 = vertex2;
    }

    public double Length()
    {
      double dx = Point2.X - Point1.X;
      double dy = Point2.Y - Point1.Y;

      return Math.Sqrt(dx * dx + dy * dy);
    }
  }

  public static class EditorExtension
  {
    public static void Zoom(this Editor ed, Extents3d ext)
    {
      if (ed == null)
        throw new ArgumentNullException("ed");
      using (ViewTableRecord view = ed.GetCurrentView())
      {
        Matrix3d worldToEye =
          Matrix3d.WorldToPlane(view.ViewDirection)
          * Matrix3d.Displacement(Point3d.Origin - view.Target)
          * Matrix3d.Rotation(view.ViewTwist, view.ViewDirection, view.Target);
        ext.TransformBy(worldToEye);
        view.Width = ext.MaxPoint.X - ext.MinPoint.X;
        view.Height = ext.MaxPoint.Y - ext.MinPoint.Y;
        view.CenterPoint = new Point2d(
          (ext.MaxPoint.X + ext.MinPoint.X) / 2.0,
          (ext.MaxPoint.Y + ext.MinPoint.Y) / 2.0
        );
        ed.SetCurrentView(view);
      }
    }

    public static void ZoomExtents(this Editor ed)
    {
      Database db = ed.Document.Database;
      db.UpdateExt(false);
      Extents3d ext =
        (short)Application.GetSystemVariable("cvport") == 1
          ? new Extents3d(db.Pextmin, db.Pextmax)
          : new Extents3d(db.Extmin, db.Extmax);
      ed.Zoom(ext);
    }
  }

  public class MagentaObject
  {
    public List<Point3d> BoundaryPoints { get; set; }
    public Point3d CenterNode { get; set; }
    public Point3d MinPoint { get; set; }
    public Point3d MaxPoint { get; set; }
    public List<Point3d> MidpointsBetweenBoundaryPoints { get; set; }

    public MagentaObject(List<Point3d> boundaryPoints)
    {
      BoundaryPoints = boundaryPoints;
      CalculateCenterPoint();
      CalculateMinMaxPoints();
      CalculateMidpointsBetweenBoundaryPoints();
    }

    private void CalculateMidpointsBetweenBoundaryPoints()
    {
      MidpointsBetweenBoundaryPoints = new List<Point3d>();

      for (int i = 0; i < BoundaryPoints.Count - 1; i++)
      {
        Point3d point3 = BoundaryPoints[i];
        Point3d point4 = BoundaryPoints[i + 1];
        Point3d midpoint1 = new Point3d((point3.X + point4.X) / 2, (point3.Y + point4.Y) / 2, 0);
        MidpointsBetweenBoundaryPoints.Add(midpoint1);
      }

      Point3d point1 = BoundaryPoints[BoundaryPoints.Count - 1];
      Point3d point2 = BoundaryPoints[0];
      Point3d midpoint = new Point3d((point1.X + point2.X) / 2, (point1.Y + point2.Y) / 2, 0);
      MidpointsBetweenBoundaryPoints.Add(midpoint);
    }

    private void CalculateCenterPoint()
    {
      double sumX = 0;
      double sumY = 0;
      foreach (Point3d point in BoundaryPoints)
      {
        sumX += point.X;
        sumY += point.Y;
      }
      double centerX = sumX / BoundaryPoints.Count;
      double centerY = sumY / BoundaryPoints.Count;
      CenterNode = new Point3d(centerX, centerY, 0);
    }

    private void CalculateMinMaxPoints()
    {
      double minX = double.MaxValue;
      double minY = double.MaxValue;
      double maxX = double.MinValue;
      double maxY = double.MinValue;

      foreach (Point3d point in BoundaryPoints)
      {
        minX = Math.Min(minX, point.X);
        minY = Math.Min(minY, point.Y);
        maxX = Math.Max(maxX, point.X);
        maxY = Math.Max(maxY, point.Y);
      }

      MinPoint = new Point3d(minX, minY, 0);
      MaxPoint = new Point3d(maxX, maxY, 0);
    }

    public TriangleNet.Geometry.Vertex CenterPointAsVertex()
    {
      return new TriangleNet.Geometry.Vertex(CenterNode.X, CenterNode.Y, 0);
    }
  }
}
