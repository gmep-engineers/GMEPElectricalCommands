using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Accord.Math;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using DocumentFormat.OpenXml.Presentation;
//using DocumentFormat.OpenXml.Wordprocessing;
using Emgu.CV.Dnn;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml.Drawing;
using Org.BouncyCastle.Bcpg.OpenPgp;
using TriangleNet.Tools;

namespace ElectricalCommands
{
  public class CADObjectCommands
  {
    public static double Scale { get; set; } = -1.0;

    public static double Voltage { get; set; } = -1.0;
    public static double MaxVoltageDropPercent { get; set; } = -1.0;
    public static int Phase { get; set; } = -1;

    public static string Address = "";

    public static string equipId,
      parentId,
      equipNo;

    public static Point3d firstClickPoint;

    public static bool HorizontalConduit = false;

    public static Point3d PanelLocation { get; set; } = new Point3d(0, 0, 0);

    public static bool IsInModel()
    {
      if (Application.DocumentManager.MdiActiveDocument.Database.TileMode)
        return true;
      else
        return false;
    }

    public static bool IsInLayout()
    {
      return !IsInModel();
    }

    public static bool IsInLayoutPaper()
    {
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      if (db.TileMode)
        return false;
      else
      {
        if (db.PaperSpaceVportId == ObjectId.Null)
          return false;
        else if (ed.CurrentViewportObjectId == ObjectId.Null)
          return false;
        else if (ed.CurrentViewportObjectId == db.PaperSpaceVportId)
          return true;
        else
          return false;
      }
    }

    public static bool IsInLayoutViewport()
    {
      return IsInLayout() && !IsInLayoutPaper();
    }

    [CommandMethod("StoreBlockData")]
    public void StoreBlockData()
    {
      if (Scale == -1.0)
      {
        SetScale();
      }

      if (Scale == -1.0)
      {
        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(
          "Please set the scale using the SetScale command before creating objects."
        );
        return;
      }

      var data = new ObjectData();

      Autodesk.AutoCAD.EditorInput.Editor ed = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument
        .Editor;
      Autodesk.AutoCAD.EditorInput.PromptSelectionResult selectionResult = ed.GetSelection();
      if (selectionResult.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
      {
        Autodesk.AutoCAD.EditorInput.SelectionSet selectionSet = selectionResult.Value;
        Autodesk.AutoCAD.EditorInput.PromptPointOptions originOptions =
          new Autodesk.AutoCAD.EditorInput.PromptPointOptions("Select an origin point: ");
        Autodesk.AutoCAD.EditorInput.PromptPointResult originResult = ed.GetPoint(originOptions);
        if (originResult.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
        {
          Point3d origin = originResult.Value;

          foreach (
            Autodesk.AutoCAD.DatabaseServices.ObjectId objectId in selectionSet.GetObjectIds()
          )
          {
            using (
              Transaction transaction = objectId.Database.TransactionManager.StartTransaction()
            )
            {
              Autodesk.AutoCAD.DatabaseServices.DBObject obj = transaction.GetObject(
                objectId,
                Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead
              );

              if (obj is Autodesk.AutoCAD.DatabaseServices.Polyline)
              {
                data = HandlePolyline(
                  obj as Autodesk.AutoCAD.DatabaseServices.Polyline,
                  data,
                  origin
                );
              }
              else if (obj is Autodesk.AutoCAD.DatabaseServices.Polyline2d)
              {
                data = HandlePolyline2d(
                  obj as Autodesk.AutoCAD.DatabaseServices.Polyline2d,
                  data,
                  origin
                );
              }
              else if (obj is Autodesk.AutoCAD.DatabaseServices.Arc)
              {
                data = HandleArc(obj as Autodesk.AutoCAD.DatabaseServices.Arc, data, origin);
              }
              else if (obj is Autodesk.AutoCAD.DatabaseServices.Circle)
              {
                data = HandleCircle(obj as Autodesk.AutoCAD.DatabaseServices.Circle, data, origin);
              }
              else if (obj is Autodesk.AutoCAD.DatabaseServices.Ellipse)
              {
                data = HandleEllipse(
                  obj as Autodesk.AutoCAD.DatabaseServices.Ellipse,
                  data,
                  origin
                );
              }
              else if (obj is Autodesk.AutoCAD.DatabaseServices.MText)
              {
                data = HandleMText(obj as Autodesk.AutoCAD.DatabaseServices.MText, data, origin);
              }
              else if (obj is Autodesk.AutoCAD.DatabaseServices.Solid)
              {
                data = HandleSolid(obj as Autodesk.AutoCAD.DatabaseServices.Solid, data, origin);
              }
              else if (obj is Autodesk.AutoCAD.DatabaseServices.Line)
              {
                data = HandleLine(obj as Autodesk.AutoCAD.DatabaseServices.Line, data, origin);
              }
              else if (obj is Autodesk.AutoCAD.DatabaseServices.DBText)
              {
                data = HandleText(obj as Autodesk.AutoCAD.DatabaseServices.DBText, data, origin);
              }
              else if (obj is Autodesk.AutoCAD.DatabaseServices.AttributeDefinition)
              {
                data = HandleAttributeDefinition(
                  obj as Autodesk.AutoCAD.DatabaseServices.AttributeDefinition,
                  data,
                  origin
                );
              }
              else if (obj is Autodesk.AutoCAD.DatabaseServices.Hatch)
              {
                data = HandleHatch(obj as Autodesk.AutoCAD.DatabaseServices.Hatch, data, origin);
              }

              transaction.Commit();
            }
          }
        }
      }

      // Prompt the user to enter a name for the JSON file
      Autodesk.AutoCAD.EditorInput.PromptStringOptions nameOptions =
        new Autodesk.AutoCAD.EditorInput.PromptStringOptions("Enter a name for the JSON file: ");

      Autodesk.AutoCAD.EditorInput.PromptResult nameResult = ed.GetString(nameOptions);

      if (nameResult.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
      {
        string fileName = nameResult.StringResult;

        // Get the directory path of the assembly
        string assemblyDirectory = System.IO.Path.GetDirectoryName(
          System.Reflection.Assembly.GetExecutingAssembly().Location
        );

        // Go up two directories from the assembly directory to get the project directory
        string projectDirectory = System
          .IO.Directory.GetParent(System.IO.Directory.GetParent(assemblyDirectory).FullName)
          .FullName;

        // Create the BlockData directory if it doesn't exist
        string blockDataDirectory = System.IO.Path.Combine(projectDirectory, "BlockData");
        if (!System.IO.Directory.Exists(blockDataDirectory))
        {
          System.IO.Directory.CreateDirectory(blockDataDirectory);
        }

        // Generate the JSON file path
        string jsonFilePath = System.IO.Path.Combine(blockDataDirectory, $"{fileName}.json");

        // Check if the file already exists
        if (System.IO.File.Exists(jsonFilePath))
        {
          // Prompt the user to confirm overwriting the existing file
          var confirmOptions = new Autodesk.AutoCAD.EditorInput.PromptStringOptions(
            "\nThe file already exists. Do you want to overwrite it? [Y/N] "
          );
          confirmOptions.AllowSpaces = false;
          var confirmResult = ed.GetString(confirmOptions);

          if (
            confirmResult.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK
            && (
              confirmResult.StringResult.Equals("Y", StringComparison.OrdinalIgnoreCase)
              || confirmResult.StringResult.Equals("Yes", StringComparison.OrdinalIgnoreCase)
            )
          )
          {
            // Save the object data to the JSON file, overwriting the existing file
            SaveDataToJsonFile(data, jsonFilePath);
          }
          else
          {
            ed.WriteMessage("\nFile not overwritten.");
          }
        }
        else
        {
          // Save the object data to the JSON file
          SaveDataToJsonFile(data, jsonFilePath);
        }
      }
    }

    [CommandMethod("SetAddress")]
    public static void SetAddress()
    {
      var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      var ed = doc.Editor;

      var promptStringOptions = new PromptStringOptions(
        "\nEnter the project address (do not paste): "
      );
      promptStringOptions.AllowSpaces = true;
      var promptStringResult = ed.GetString(promptStringOptions);
      Address = promptStringResult.StringResult.ToUpper();
    }

    [CommandMethod("SetScale")]
    public static void SetScale()
    {
      var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      var ed = doc.Editor;

      var promptStringOptions = new PromptStringOptions(
        "\nEnter the scale value (e.g., 1/4, 3/16, 1/8): "
      );
      var promptStringResult = ed.GetString(promptStringOptions);

      if (promptStringResult.Status == PromptStatus.OK)
      {
        string scaleString = promptStringResult.StringResult;
        string[] scaleParts = scaleString.Split('/');

        if (
          scaleParts.Length == 2
          && double.TryParse(scaleParts[0], out double numerator)
          && double.TryParse(scaleParts[1], out double denominator)
        )
        {
          Scale = numerator / denominator;
          ed.WriteMessage($"\nScale set to {scaleString} ({Scale})");
        }
        else
        {
          ed.WriteMessage(
            $"\nInvalid scale format. Please enter the scale in the format 'numerator/denominator' (e.g., 1/4, 3/16, 1/8)."
          );
        }
      }
    }

    [CommandMethod("SetPanelLocation")]
    public static void SetPanelLocation()
    {
      var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      var ed = doc.Editor;

      var promptPointOptions = new Autodesk.AutoCAD.EditorInput.PromptPointOptions(
        "\nSelect the panel location: "
      );
      var promptPointResult = ed.GetPoint(promptPointOptions);

      if (promptPointResult.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
      {
        PanelLocation = promptPointResult.Value;
        ed.WriteMessage($"\nPanel location set to {PanelLocation}");
      }
      else
      {
        ed.WriteMessage("\nInvalid point. Panel location not set.");
      }
    }

    [CommandMethod("CreateAllSymbols")]
    public static void CreateAllSymbols()
    {
      if (Scale == -1.0)
      {
        SetScale();
      }

      if (Scale == -1.0)
      {
        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(
          "Please set the scale using the SetScale command before creating objects."
        );
        return;
      }

      Document acDoc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database acCurDb = acDoc.Database;
      Editor ed = acDoc.Editor;

      // Prompt the user to select a start location
      PromptPointResult ppr = ed.GetPoint("\nSpecify start location for symbols: ");
      if (ppr.Status != PromptStatus.OK)
      {
        ed.WriteMessage("\nOperation canceled.");
        return;
      }
      Point3d startPoint = ppr.Value;

      string assemblyDirectory = System.IO.Path.GetDirectoryName(
        System.Reflection.Assembly.GetExecutingAssembly().Location
      );
      string projectDirectory = System
        .IO.Directory.GetParent(System.IO.Directory.GetParent(assemblyDirectory).FullName)
        .FullName;
      string blockDataDirectory = System.IO.Path.Combine(projectDirectory, "BlockData");

      if (System.IO.Directory.Exists(blockDataDirectory))
      {
        string[] jsonFiles = System.IO.Directory.GetFiles(blockDataDirectory, "*.json");

        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        {
          BlockTable acBlkTbl =
            acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
          BlockTableRecord acBlkTblRec =
            acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite)
            as BlockTableRecord;

          Point3d currentInsertionPoint = startPoint;

          foreach (string jsonFile in jsonFiles)
          {
            string jsonData = System.IO.File.ReadAllText(jsonFile);
            ObjectData objectData = JsonConvert.DeserializeObject<ObjectData>(jsonData);
            double spacing = 8 / Scale; // Adjust spacing based on scale

            CreateObjectFromData(jsonData, currentInsertionPoint, acBlkTblRec);

            // Add text object above each symbol using filename without .json extension
            string fileName = System.IO.Path.GetFileNameWithoutExtension(jsonFile);
            Point3d textPosition = new Point3d(
              currentInsertionPoint.X,
              currentInsertionPoint.Y + (10 / Scale),
              currentInsertionPoint.Z
            ); // Adjust text position based on scale
            DBText text = new DBText
            {
              Position = textPosition,
              Height = 1.125 / Scale,
              TextString = fileName,
              Layer = "E-TXT1",
              HorizontalMode = TextHorizontalMode.TextCenter,
              AlignmentPoint = textPosition,
            };
            text.HorizontalMode = TextHorizontalMode.TextCenter;
            text.AlignmentPoint = textPosition;
            acBlkTblRec.AppendEntity(text);
            acTrans.AddNewlyCreatedDBObject(text, true);

            currentInsertionPoint = new Point3d(
              currentInsertionPoint.X + spacing,
              currentInsertionPoint.Y,
              currentInsertionPoint.Z
            );
          }

          acTrans.Commit();
        }

        ed.WriteMessage($"\nCreated {jsonFiles.Length} symbols from the BlockData directory.");
      }
      else
      {
        ed.WriteMessage("\nBlockData directory not found.");
      }
    }

    [CommandMethod("TXL")]
    public void TXL()
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Editor ed = doc.Editor;

      PromptPointOptions ppo = new PromptPointOptions("\nSelect start point:");
      PromptPointResult ppr = ed.GetPoint(ppo);
      if (ppr.Status != PromptStatus.OK)
        return;

      CreateMText(ppr.Value, AttachmentPoint.TopLeft);
    }

    [CommandMethod("TXR")]
    public void TXR()
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Editor ed = doc.Editor;

      PromptPointOptions ppo = new PromptPointOptions("\nSelect start point:");
      PromptPointResult ppr = ed.GetPoint(ppo);
      if (ppr.Status != PromptStatus.OK)
        return;

      CreateMText(ppr.Value, AttachmentPoint.TopRight);
    }

    [CommandMethod("TXC")]
    public void TXC()
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Editor ed = doc.Editor;

      PromptPointOptions ppo = new PromptPointOptions("\nSelect start point:");
      PromptPointResult ppr = ed.GetPoint(ppo);
      if (ppr.Status != PromptStatus.OK)
        return;

      CreateMText(ppr.Value, AttachmentPoint.TopCenter);
    }

    [CommandMethod("GFI")]
    public static void GFI()
    {
      CreateTextWithJig("E-TXT3", TextHorizontalMode.TextLeft, "GFI");
    }

    [CommandMethod("CT")]
    public static void CT()
    {
      CreateTextWithJig("E-TXT3", TextHorizontalMode.TextLeft, "+48\"");
    }

    [CommandMethod("AR")]
    public static void AR()
    {
      if (Scale == -1.0)
      {
        SetScale();
      }

      if (Scale == -1.0)
      {
        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(
          "Please set the scale using the SetScale command before creating objects."
        );
        return;
      }

      if (PanelLocation == new Point3d(0, 0, 0))
      {
        SetPanelLocation();
      }

      Document acDoc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database acCurDb = acDoc.Database;
      Editor ed = acDoc.Editor;

      ed.WriteMessage($"\nCurrent panel location is set to {PanelLocation}");
      ed.WriteMessage("\nRun the 'SetPanelLocation' command to set a new panel location.");

      string assemblyDirectory = System.IO.Path.GetDirectoryName(
        System.Reflection.Assembly.GetExecutingAssembly().Location
      );
      string jsonFilePath = System.IO.Path.Combine(
        assemblyDirectory,
        "Symbols",
        "BlockData",
        "ar.json"
      );

      if (File.Exists(jsonFilePath))
      {
        string jsonData = File.ReadAllText(jsonFilePath);

        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        {
          BlockTable acBlkTbl =
            acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
          BlockTableRecord acBlkTblRec;

          if (acBlkTbl.Has($"ar{Scale}"))
          {
            acBlkTblRec =
              acTrans.GetObject(acBlkTbl[$"ar{Scale}"], OpenMode.ForWrite) as BlockTableRecord;
            foreach (ObjectId id in acBlkTblRec)
            {
              acTrans.GetObject(id, OpenMode.ForWrite).Erase();
            }
          }
          else
          {
            acBlkTblRec = new BlockTableRecord();
            acBlkTblRec.Name = $"ar{Scale}";

            acBlkTbl.UpgradeOpen();
            acBlkTbl.Add(acBlkTblRec);
            acTrans.AddNewlyCreatedDBObject(acBlkTblRec, true);
          }

          CreateObjectFromData(jsonData, Point3d.Origin, acBlkTblRec);

          acTrans.Commit();
        }

        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        {
          BlockTable acBlkTbl =
            acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
          BlockTableRecord acBlkTblRec =
            acTrans.GetObject(acBlkTbl[$"ar{Scale}"], OpenMode.ForRead) as BlockTableRecord;

          using (BlockReference acBlkRef = new BlockReference(Point3d.Origin, acBlkTblRec.ObjectId))
          {
            ArrowJig arrowJig = new ArrowJig(acBlkRef, PanelLocation);
            PromptResult promptResult = ed.Drag(arrowJig);

            if (promptResult.Status == PromptStatus.OK)
            {
              BlockTableRecord currentSpace =
                acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
              currentSpace.AppendEntity(acBlkRef);
              acTrans.AddNewlyCreatedDBObject(acBlkRef, true);

              // Set block reference layer
              acBlkRef.Layer = "E-CND1";

              acTrans.Commit();

              Point3d insertionPoint = arrowJig.InsertionPoint;

              PromptPointResult ppr2 = ed.GetPoint("\nSpecify second point of the spline: ");
              if (ppr2.Status == PromptStatus.OK)
              {
                Point3d secondPoint = ppr2.Value;

                PromptPointResult ppr3 = ed.GetPoint("\nSpecify third point of the spline: ");
                if (ppr3.Status == PromptStatus.OK)
                {
                  Point3d thirdPoint = ppr3.Value;

                  using (Transaction acTransSpline = acCurDb.TransactionManager.StartTransaction())
                  {
                    currentSpace =
                      acTransSpline.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite)
                      as BlockTableRecord;

                    using (
                      Spline acSpline = new Spline(
                        new Point3dCollection(new[] { insertionPoint, secondPoint, thirdPoint }),
                        0,
                        0.0
                      )
                    )
                    {
                      acSpline.Layer = "E-CND1";
                      currentSpace.AppendEntity(acSpline);
                      acTransSpline.AddNewlyCreatedDBObject(acSpline, true);
                    }

                    acTransSpline.Commit();
                  }
                }
              }
            }
          }
        }
      }
      else
      {
        ed.WriteMessage($"\nBlock 'ar' not found in BlockData directory.");
      }
    }

    public List<Dictionary<string, object>> GetPanelJsonData()
    {
      List<Dictionary<string, object>> allPanelData = new List<Dictionary<string, object>>();
      Document acDoc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      string acDocPath = Path.GetDirectoryName(acDoc.Name);
      string savesDirectory = Path.Combine(acDocPath, "Saves");
      string panelSavesDirectory = Path.Combine(savesDirectory, "Panel");

      // Check if the "Saves/Panel" directory exists
      if (Directory.Exists(panelSavesDirectory))
      {
        // Get all JSON files in the directory
        string[] jsonFiles = Directory.GetFiles(panelSavesDirectory, "*.json");

        // If there are any JSON files, find the most recent one
        if (jsonFiles.Length > 0)
        {
          string mostRecentJsonFile = jsonFiles
            .OrderByDescending(f => File.GetLastWriteTime(f))
            .First();

          // Read the JSON data from the file
          string jsonData = File.ReadAllText(mostRecentJsonFile);

          // Deserialize the JSON data to a list of dictionaries
          allPanelData = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(jsonData);
        }
      }

      return allPanelData;
    }

    [CommandMethod("PANELLOAD")]
    public void LINKPANELKVA()
    {
      var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      var db = doc.Database;
      var ed = doc.Editor;
      var linkPanelNamePrompt = new PromptStringOptions("\nEnter the panel name: ");
      linkPanelNamePrompt.AllowSpaces = true;
      var linkPanelNameResult = ed.GetString(linkPanelNamePrompt);
      string panelName = "";
      while (panelName.Length == 0)
      {
        if (linkPanelNameResult.Status == PromptStatus.OK)
        {
          panelName = linkPanelNameResult
            .StringResult.ToUpper()
            .Replace("panel", "")
            .Replace(" ", "");
          List<Dictionary<string, object>> panelData = GetPanelJsonData();
          string panelId = "";
          foreach (Dictionary<string, object> data in panelData)
          {
            string n = data["panel"] as string;
            if (n.Replace(" ", "").Replace("'", "") == panelName)
            {
              panelId = data["id"] as string;
              panelId = panelId.Replace("-", "");
            }
          }
          if (String.IsNullOrEmpty(panelId))
          {
            return;
          }
          string panelKvaString = (string)doc.GetLispSymbol($"panel_{panelId}_kva");
          string panelAString = (string)doc.GetLispSymbol($"panel_{panelId}_a");
          if (panelKvaString != null)
          {
            PromptPointOptions ppo = new PromptPointOptions("\nSpecify insertion point: ");
            PromptPointResult ppr = ed.GetPoint(ppo);
            if (ppr.Status != PromptStatus.OK)
              return;
            using (Transaction t = db.TransactionManager.StartTransaction())
            {
              BlockTable acBlkTbl = t.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
              BlockTableRecord acBlkTblRec =
                t.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite)
                as BlockTableRecord;
              var textStyleId = GeneralCommands.GetTextStyleId("gmep");
              var textStyle = (TextStyleTableRecord)t.GetObject(textStyleId, OpenMode.ForRead);
              List<DBText> allTexts = new List<DBText>();
              double scaleFactor = 1;
              if (IsInModel() || IsInLayoutViewport())
              {
                if (Scale <= 0)
                {
                  SetScale();
                }
                scaleFactor = 12 / Scale;
              }
              var kvaText = new DBText
              {
                Position = new Point3d(ppr.Value.X, ppr.Value.Y - 0.16 * scaleFactor, 0),
                Height = 0.1 * scaleFactor,
                WidthFactor = 0.85,
                Layer = "E-TXT1",
                TextStyleId = textStyleId,
                HorizontalMode = TextHorizontalMode.TextLeft,
                VerticalMode = TextVerticalMode.TextVerticalMid,
                Justify = AttachmentPoint.BaseLeft,
              };
              var aText = new DBText
              {
                Position = new Point3d(ppr.Value.X, ppr.Value.Y - 0.32 * scaleFactor, 0),
                Height = 0.1 * scaleFactor,
                WidthFactor = 0.85,
                Layer = "E-TXT1",
                TextStyleId = textStyleId,
                HorizontalMode = TextHorizontalMode.TextLeft,
                VerticalMode = TextVerticalMode.TextVerticalMid,
                Justify = AttachmentPoint.BaseLeft,
              };
              acBlkTblRec.AppendEntity(kvaText);
              acBlkTblRec.AppendEntity(aText);
              t.AddNewlyCreatedDBObject(kvaText, true);
              t.AddNewlyCreatedDBObject(aText, true);
              string kvaFieldFormat = $"%<\\AcVar.17.0 Lisp.panel_{panelId}_kva>%";
              string aFieldFormat = $"%<\\AcVar.17.0 Lisp.panel_{panelId}_a>%";
              Field kvaField = new Field(kvaFieldFormat);
              Field aField = new Field(aFieldFormat);
              kvaField.Evaluate();
              aField.Evaluate();
              kvaText.SetField(kvaField);
              aText.SetField(aField);
              t.AddNewlyCreatedDBObject(kvaField, true);
              t.AddNewlyCreatedDBObject(aField, true);
              t.Commit();
            }
          }
          else
          {
            return;
          }
        }
        else
        {
          return;
        }
      }
    }

    [CommandMethod("SetVoltage")]
    public static void SetVoltage()
    {
      Voltage = 0;
      var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      var ed = doc.Editor;

      var voltagePrompt = new PromptStringOptions("\nEnter the voltage: ");
      var voltageResult = ed.GetString(voltagePrompt);
      string validVoltages = "115;120;208;230;240;277;460;480";
      while (Voltage <= 0)
      {
        if (voltageResult.Status == PromptStatus.OK)
        {
          string voltageString = voltageResult.StringResult;
          if (validVoltages.Contains(voltageString) && double.TryParse(voltageString, out double v))
          {
            Voltage = v;
            ed.WriteMessage($"\nVoltage set to {Voltage}");
          }
          else
          {
            ed.WriteMessage($"\nInvalid voltage.");
            voltageResult = ed.GetString(voltagePrompt);
          }
        }
        else
        {
          return;
        }
      }
    }

    [CommandMethod("SetMaxVoltageDropPercent")]
    public static void SetMaxVoltageDropPercent()
    {
      MaxVoltageDropPercent = 0;
      var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      var ed = doc.Editor;

      var voltageDropPrompt = new PromptStringOptions("\nEnter the max voltage drop percent: ");
      var voltageDropResult = ed.GetString(voltageDropPrompt);
      while (MaxVoltageDropPercent <= 0)
      {
        if (voltageDropResult.Status == PromptStatus.OK)
        {
          string voltageDropString = voltageDropResult.StringResult;

          if (double.TryParse(voltageDropString, out double vd))
          {
            MaxVoltageDropPercent = vd;
            ed.WriteMessage($"\nMax voltage drop set to {MaxVoltageDropPercent}");
          }
          else
          {
            ed.WriteMessage($"\nInvalid percent.");
            voltageDropResult = ed.GetString(voltageDropPrompt);
          }
        }
        else
        {
          return;
        }
      }
    }

    [CommandMethod("SetPhase")]
    public static void SetPhase()
    {
      Phase = 0;
      var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      var ed = doc.Editor;

      var phasePrompt = new PromptStringOptions("\nEnter the phase: ");
      var phaseResult = ed.GetString(phasePrompt);
      while (Phase <= 0)
      {
        if (phaseResult.Status == PromptStatus.OK)
        {
          string phaseString = phaseResult.StringResult;
          if ((phaseString == "1" || phaseString == "3") && int.TryParse(phaseString, out int ph))
          {
            Phase = ph;
            ed.WriteMessage($"\nPhase set to {Phase}");
          }
          else
          {
            ed.WriteMessage($"\nInvalid phase.");
            phaseResult = ed.GetString(phasePrompt);
          }
        }
        else
        {
          return;
        }
      }
    }

    public struct WireSpec
    {
      public int parallelWires;
      public string wireSize;
      public double actualVoltageDrop;
      public int wireSizeIndex;
    }

    public struct ConduitSpec
    {
      public WireSpec wireSpec;
      public string conduitSize;
    }

    private static double GetVoltageDrop(
      string wireSize,
      double distance,
      int parallelWires,
      double loadAmperage,
      double multiplier
    )
    {
      double factor = 1.0 / parallelWires * distance * loadAmperage * multiplier;
      switch (wireSize)
      {
        case "12":
          return 0.0020500 * factor;
        case "10":
          return 0.0012900 * factor;
        case "8":
          return 0.0008090 * factor;
        case "6":
          return 0.0005100 * factor;
        case "4":
          return 0.0003210 * factor;
        case "3":
          return 0.0002540 * factor;
        case "2":
          return 0.0002010 * factor;
        case "1":
          return 0.0001600 * factor;
        case "1/0":
          return 0.0001270 * factor;
        case "2/0":
          return 0.0001010 * factor;
        case "3/0":
          return 0.0000797 * factor;
        case "4/0":
          return 0.0000626 * factor;
        case "250 KCMIL":
          return 0.0000535 * factor;
        case "300 KCMIL":
          return 0.0000446 * factor;
        case "350 KCMIL":
          return 0.0000382 * factor;
        case "400 KCMIL":
          return 0.0000331 * factor;
        case "500 KCMIL":
          return 0.0000265 * factor;
      }
      return -1;
    }

    private static WireSpec GetWireSize(
      double amperage,
      double distance,
      double multiplier,
      double maxVoltageDropAllowed,
      int parallelWires = 1
    )
    {
      int wireSizeIndex = 0;
      double actualVoltageDrop = 600;
      Dictionary<string, double> resistancePerFoot = new Dictionary<string, double>();
      resistancePerFoot.Add("12", 0.0020500);
      resistancePerFoot.Add("10", 0.0012900);
      resistancePerFoot.Add("8", 0.0008090);
      resistancePerFoot.Add("6", 0.0005100);
      resistancePerFoot.Add("4", 0.0003210);
      resistancePerFoot.Add("3", 0.0002540);
      resistancePerFoot.Add("2", 0.0002010);
      resistancePerFoot.Add("1", 0.0001600);
      resistancePerFoot.Add("1/0", 0.0001270);
      resistancePerFoot.Add("2/0", 0.0001010);
      resistancePerFoot.Add("3/0", 0.0000797);
      resistancePerFoot.Add("4/0", 0.0000626);
      resistancePerFoot.Add("250 KCMIL", 0.0000535);
      resistancePerFoot.Add("300 KCMIL", 0.0000446);
      resistancePerFoot.Add("350 KCMIL", 0.0000382);
      resistancePerFoot.Add("400 KCMIL", 0.0000331);
      resistancePerFoot.Add("500 KCMIL", 0.0000265);
      while (actualVoltageDrop > maxVoltageDropAllowed)
      {
        double resistance = resistancePerFoot.ElementAt(wireSizeIndex).Value;
        actualVoltageDrop = resistance / parallelWires * amperage * distance * multiplier;
        if (actualVoltageDrop > maxVoltageDropAllowed)
        {
          wireSizeIndex++;
        }
        if (wireSizeIndex == resistancePerFoot.Count && actualVoltageDrop > maxVoltageDropAllowed)
        {
          wireSizeIndex = 0;
          parallelWires++;
        }
      }
      WireSpec wireSpec = new WireSpec();
      wireSpec.parallelWires = parallelWires;
      wireSpec.wireSize = resistancePerFoot.ElementAt(wireSizeIndex).Key;
      wireSpec.actualVoltageDrop = actualVoltageDrop;
      wireSpec.wireSizeIndex = wireSizeIndex;
      double maxWireAmpacity = GetMaxWireAmpacity(wireSpec.wireSize);
      while (maxWireAmpacity < Math.Round(amperage / parallelWires, 0))
      {
        if (wireSizeIndex == resistancePerFoot.Count - 1)
        {
          wireSizeIndex = 0;
          parallelWires++;
        }
        else
        {
          wireSizeIndex++;
        }
        wireSpec.wireSize = resistancePerFoot.ElementAt(wireSizeIndex).Key;
        wireSpec.parallelWires = parallelWires;
        double resistance = resistancePerFoot.ElementAt(wireSizeIndex).Value;
        wireSpec.actualVoltageDrop = resistance / parallelWires * amperage * distance * multiplier;
        wireSpec.wireSizeIndex = wireSizeIndex;
        maxWireAmpacity = GetMaxWireAmpacity(wireSpec.wireSize);
      }
      return wireSpec;
    }

    private static ConduitSpec GetConduitAndWireSize(
      double loadAmperage,
      double mocp,
      double distance,
      double multiplier,
      double maxVoltageDropAllowed,
      int wires
    )
    {
      WireSpec maxWireSpec = GetWireSize(mocp, distance, multiplier, maxVoltageDropAllowed);
      WireSpec loadWireSpec = GetWireSize(
        loadAmperage,
        distance,
        multiplier,
        maxVoltageDropAllowed,
        maxWireSpec.parallelWires
      );
      WireSpec minWireSpecPerMocp = GetWireSize(
        mocp,
        1,
        multiplier,
        maxVoltageDropAllowed,
        maxWireSpec.parallelWires
      );
      if (loadWireSpec.wireSizeIndex < minWireSpecPerMocp.wireSizeIndex)
      {
        loadWireSpec = minWireSpecPerMocp;
      }
      ConduitSpec spec = new ConduitSpec();
      if (wires == 4)
      {
        Dictionary<string, string> conduitSize4W = new Dictionary<string, string>();
        conduitSize4W.Add("12", "3/4");
        conduitSize4W.Add("10", "3/4");
        conduitSize4W.Add("8", "3/4");
        conduitSize4W.Add("6", "3/4");
        conduitSize4W.Add("4", "1");
        conduitSize4W.Add("3", "1-1/4");
        conduitSize4W.Add("2", "1-1/4");
        conduitSize4W.Add("1", "1-1/4");
        conduitSize4W.Add("1/0", "1-1/2");
        conduitSize4W.Add("2/0", "2");
        conduitSize4W.Add("3/0", "2");
        conduitSize4W.Add("4/0", "2");
        conduitSize4W.Add("250 KCMIL", "2-1/2");
        conduitSize4W.Add("300 KCMIL", "2-1/2");
        conduitSize4W.Add("350 KCMIL", "3");
        conduitSize4W.Add("400 KCMIL", "3");
        conduitSize4W.Add("500 KCMIL", "3");
        if (conduitSize4W.TryGetValue(maxWireSpec.wireSize, out string size))
        {
          spec.wireSpec = loadWireSpec;
          spec.conduitSize = size;
        }
      }
      else
      {
        Dictionary<string, string> conduitSize3W = new Dictionary<string, string>();
        conduitSize3W.Add("12", "3/4");
        conduitSize3W.Add("10", "3/4");
        conduitSize3W.Add("8", "3/4");
        conduitSize3W.Add("6", "3/4");
        conduitSize3W.Add("4", "1");
        conduitSize3W.Add("3", "1");
        conduitSize3W.Add("2", "1");
        conduitSize3W.Add("1", "1-1/4");
        conduitSize3W.Add("1/0", "1-1/4");
        conduitSize3W.Add("2/0", "1-1/2");
        conduitSize3W.Add("3/0", "1-1/2");
        conduitSize3W.Add("4/0", "2");
        conduitSize3W.Add("250 KCMIL", "2");
        conduitSize3W.Add("300 KCMIL", "2");
        conduitSize3W.Add("350 KCMIL", "2-1/2");
        conduitSize3W.Add("400 KCMIL", "2-1/2");
        conduitSize3W.Add("500 KCMIL", "3");
        if (conduitSize3W.TryGetValue(maxWireSpec.wireSize, out string size))
        {
          spec.wireSpec = loadWireSpec;
          spec.conduitSize = size;
        }
      }
      return spec;
    }

    private static string GetGroundingSize(double mocp)
    {
      string gndSize = "";
      mocp = Math.Round(mocp, 0);
      switch (mocp)
      {
        case var _ when mocp <= 20:
          gndSize = "12";
          break;
        case var _ when mocp <= 30:
          gndSize = "10";
          break;
        case var _ when mocp <= 100:
          gndSize = "8";
          break;
        case var _ when mocp <= 200:
          gndSize = "6";
          break;
        case var _ when mocp <= 300:
          gndSize = "4";
          break;
        case var _ when mocp <= 400:
          gndSize = "3";
          break;
        case var _ when mocp <= 500:
          gndSize = "2";
          break;
        case var _ when mocp <= 600:
          gndSize = "1";
          break;
        case var _ when mocp <= 800:
          gndSize = "1/0";
          break;
        case var _ when mocp <= 1000:
          gndSize = "2/0";
          break;
        case var _ when mocp <= 1200:
          gndSize = "3/0";
          break;
        case var _ when mocp <= 1600:
          gndSize = "4/0";
          break;
        case var _ when mocp <= 2000:
          gndSize = "250 KCMIL";
          break;
        case var _ when mocp <= 2500:
          gndSize = "350 KCMIL";
          break;
        case var _ when mocp <= 3000:
          gndSize = "400 KCMIL";
          break;
        case var _ when mocp <= 4000:
          gndSize = "500 KCMIL";
          break;
        case var _ when mocp <= 5000:
          gndSize = "700 KCMIL";
          break;
        case var _ when mocp <= 6000:
          gndSize = "800 KCMIL";
          break;
      }
      return gndSize;
    }

    private static double GetMaxWireAmpacity(string wireSize)
    {
      switch (wireSize)
      {
        case "12":
          return 20;
        case "10":
          return 30;
        case "8":
          return 40;
        case "6":
          return 55;
        case "4":
          return 70;
        case "3":
          return 85;
        case "2":
          return 95;
        case "1":
          return 110;
        case "1/0":
          return 150;
        case "2/0":
          return 175;
        case "3/0":
          return 200;
        case "4/0":
          return 230;
        case "250 KCMIL":
          return 255;
        case "300 KCMIL":
          return 285;
        case "350 KCMIL":
          return 310;
        case "400 KCMIL":
          return 335;
        default:
          return 380;
      }
    }

    public static int GetMcaFromFla(double fla)
    {
      double mocp = fla / 0.8;
      switch (fla)
      {
        case var _ when mocp <= 20:
          return 20;
        case var _ when mocp <= 25:
          return 25;
        case var _ when mocp <= 30:
          return 30;
        case var _ when mocp <= 35:
          return 35;
        case var _ when mocp <= 40:
          return 40;
        case var _ when mocp <= 45:
          return 45;
        case var _ when mocp <= 50:
          return 50;
        case var _ when mocp <= 60:
          return 60;
        case var _ when mocp <= 70:
          return 70;
        case var _ when mocp <= 80:
          return 80;
        case var _ when mocp <= 90:
          return 90;
        case var _ when mocp <= 100:
          return 100;
        case var _ when mocp <= 110:
          return 110;
        case var _ when mocp <= 125:
          return 125;
        case var _ when mocp <= 150:
          return 150;
        case var _ when mocp <= 175:
          return 175;
        case var _ when mocp <= 200:
          return 200;
        case var _ when mocp <= 225:
          return 225;
        case var _ when mocp <= 250:
          return 250;
        default:
          return -1;
      }
    }

    public static string GetConnectionTypeFromFlaVoltage(double fla, int voltage)
    {
      int mca = GetMcaFromFla(fla);
      switch (fla)
      {
        case var _ when fla / 0.8 <= 15:
          if (voltage <= 120)
          {
            return "NEMA 5-15";
          }
          else
          {
            return "NEMA 6-15";
          }
        case var _ when mca <= 20:
          if (voltage <= 120)
          {
            return "NEMA 5-20";
          }
          else
          {
            return "NEMA 6-20";
          }
        case var _ when mca <= 30:
          if (voltage <= 120)
          {
            return "NEMA 5-30";
          }
          else
          {
            return "NEMA 6-30";
          }
        case var _ when mca <= 50:
          if (voltage <= 120)
          {
            return "NEMA 5-50";
          }
          else
          {
            return "NEMA 6-50";
          }
        case var _ when mca <= 60:
          return "60AS/" + mca + "AF/" + (voltage <= 120 ? "2P" : "3P");
        case var _ when mca <= 100:
          return "100AS/" + mca + "AF/" + (voltage <= 120 ? "2P" : "3P");
        case var _ when mca <= 200:
          return "200AS/" + mca + "AF/" + (voltage <= 120 ? "2P" : "3P");
        case var _ when mca <= 400:
          return "400AS/" + mca + "AF/" + (voltage <= 120 ? "2P" : "3P");
        case var _ when mca <= 600:
          return "600AS/" + mca + "AF/" + (voltage <= 120 ? "2P" : "3P");
      }
      return "V.I.F";
    }

    public static (string, string, string, string, string, string) GetWireAndConduitSizeText(
      double loadAmperage,
      double mocp,
      double distance,
      double voltage,
      double maxVoltageDropPercent,
      int phase
    )
    {
      double maxVoltageDropAllowed = voltage * maxVoltageDropPercent / 100;
      double minVoltageDropAllowedAtLoad = voltage * maxVoltageDropAllowed;
      double multiplier = 2.0;
      if (Phase == 3)
      {
        multiplier = 1.732;
      }

      int numWires = 3;
      if (phase == 3)
      {
        numWires = 4;
      }
      if (voltage <= 120)
      {
        numWires = 2;
      }

      ConduitSpec spec = GetConduitAndWireSize(
        loadAmperage,
        mocp,
        distance,
        multiplier,
        maxVoltageDropAllowed,
        numWires
      );
      string gndSize = "";
      double voltageDropPercent = 0;
      gndSize = GetGroundingSize(mocp);
      voltageDropPercent =
        GetVoltageDrop(
          spec.wireSpec.wireSize,
          distance,
          spec.wireSpec.parallelWires,
          loadAmperage,
          multiplier
        )
        / voltage
        * 100;
      string firstLine;
      string secondLine;
      string thirdLine;
      if (spec.wireSpec.parallelWires > 1)
      {
        firstLine =
          $"[{spec.wireSpec.parallelWires}]{spec.conduitSize}\" C.; {numWires}#{spec.wireSpec.wireSize} CU.";
      }
      else
      {
        firstLine = $"{spec.conduitSize}\" C.; {numWires}#{spec.wireSpec.wireSize} CU.";
      }
      secondLine = $"PLUS 1#{gndSize} CU. GND.";
      string voltageDropPercentString =
        voltageDropPercent < 0.1 ? "NEGL." : Math.Round(voltageDropPercent, 1).ToString() + "%";
      thirdLine = $"{distance}'; VD={voltageDropPercentString}";
      double loadWireSize = distance > 100 ? Math.Round(loadAmperage, 1) : mocp;
      string supplemental1 = $"C. SIZED FOR {mocp}A";
      string supplemental2 = $"W. SIZED FOR {loadWireSize}A";
      string supplemental3 = $"@{voltage}V-{phase}\u0081-{numWires}W";
      return (firstLine, secondLine, thirdLine, supplemental1, supplemental2, supplemental3);
    }

    public static void AddWireAndConduitTextToPlan(
      Database acCurDb,
      Point3d ppr,
      string firstLine,
      string secondLine,
      string thirdLine,
      string supplemental1,
      string supplemental2,
      string supplemental3,
      bool horizontal
    )
    {
      using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
      {
        BlockTable acBlkTbl =
          acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord acBlkTblRec =
          acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite)
          as BlockTableRecord;
        var textStyleId = GeneralCommands.GetTextStyleId("gmep");
        var textStyle = (TextStyleTableRecord)acTrans.GetObject(textStyleId, OpenMode.ForRead);
        List<DBText> allTexts = new List<DBText>();
        double scaleFactor = 1;
        if (IsInModel() || IsInLayoutViewport())
        {
          if (Scale <= 0)
          {
            SetScale();
          }
          scaleFactor = 12 / Scale;
        }
        var firstLineText = new DBText
        {
          TextString = firstLine,
          Position = horizontal
            ? new Point3d(ppr.X, ppr.Y + 0.20 * scaleFactor, 0)
            : new Point3d(ppr.X - 0.20 * scaleFactor, ppr.Y, 0),
          Height = 0.1 * scaleFactor,
          WidthFactor = 0.85,
          Layer = "E-TEXT",
          TextStyleId = textStyleId,
          HorizontalMode = TextHorizontalMode.TextLeft,
          VerticalMode = TextVerticalMode.TextVerticalMid,
          Justify = AttachmentPoint.BaseLeft,
          Rotation = horizontal ? 0 : 1.5708,
        };
        allTexts.Add(firstLineText);
        var secondLineText = new DBText
        {
          TextString = secondLine,
          Position = horizontal
            ? new Point3d(ppr.X, ppr.Y + 0.04 * scaleFactor, 0)
            : new Point3d(ppr.X - 0.04 * scaleFactor, ppr.Y, 0),
          Height = 0.1 * scaleFactor,
          WidthFactor = 0.85,
          Layer = "E-TEXT",
          TextStyleId = textStyleId,
          HorizontalMode = TextHorizontalMode.TextLeft,
          VerticalMode = TextVerticalMode.TextVerticalMid,
          Justify = AttachmentPoint.BaseLeft,
          Rotation = horizontal ? 0 : 1.5708,
        };
        allTexts.Add(secondLineText);
        var thirdLineText = new DBText
        {
          TextString = thirdLine,
          Position = horizontal
            ? new Point3d(ppr.X, ppr.Y - 0.13 * scaleFactor, 0)
            : new Point3d(ppr.X + 0.13 * scaleFactor, ppr.Y, 0),
          Height = 0.1 * scaleFactor,
          WidthFactor = 0.85,
          Layer = "E-TEXT",
          TextStyleId = textStyleId,
          HorizontalMode = TextHorizontalMode.TextLeft,
          VerticalMode = TextVerticalMode.TextVerticalMid,
          Justify = AttachmentPoint.BaseLeft,
          Rotation = horizontal ? 0 : 1.5708,
        };
        allTexts.Add(thirdLineText);
        var supplementalText1 = new DBText
        {
          TextString = supplemental1,
          Position = horizontal
            ? new Point3d(ppr.X, ppr.Y - 0.29 * scaleFactor, 0)
            : new Point3d(ppr.X + 0.29 * scaleFactor, ppr.Y, 0),
          Height = 0.1 * scaleFactor,
          WidthFactor = 0.85,
          Layer = "DEFPOINTS",
          TextStyleId = textStyleId,
          HorizontalMode = TextHorizontalMode.TextLeft,
          VerticalMode = TextVerticalMode.TextVerticalMid,
          Justify = AttachmentPoint.BaseLeft,
          Rotation = horizontal ? 0 : 1.5708,
        };
        allTexts.Add(supplementalText1);
        var supplementalText2 = new DBText
        {
          TextString = supplemental2,
          Position = horizontal
            ? new Point3d(ppr.X, ppr.Y - 0.45 * scaleFactor, 0)
            : new Point3d(ppr.X + 0.45 * scaleFactor, ppr.Y, 0),
          Height = 0.1 * scaleFactor,
          WidthFactor = 0.85,
          Layer = "DEFPOINTS",
          TextStyleId = textStyleId,
          HorizontalMode = TextHorizontalMode.TextLeft,
          VerticalMode = TextVerticalMode.TextVerticalMid,
          Justify = AttachmentPoint.BaseLeft,
          Rotation = horizontal ? 0 : 1.5708,
        };
        allTexts.Add(supplementalText2);
        var supplementalText3 = new DBText
        {
          TextString = supplemental3,
          Position = horizontal
            ? new Point3d(ppr.X, ppr.Y - 0.61 * scaleFactor, 0)
            : new Point3d(ppr.X + 0.61 * scaleFactor, ppr.Y, 0),
          Height = 0.1 * scaleFactor,
          WidthFactor = 0.85,
          Layer = "DEFPOINTS",
          TextStyleId = textStyleId,
          HorizontalMode = TextHorizontalMode.TextLeft,
          VerticalMode = TextVerticalMode.TextVerticalMid,
          Justify = AttachmentPoint.BaseLeft,
          Rotation = horizontal ? 0 : 1.5708,
        };
        allTexts.Add(supplementalText3);

        var currentSpace = (BlockTableRecord)
          acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
        foreach (DBText text in allTexts)
        {
          currentSpace.AppendEntity(text);
          acTrans.AddNewlyCreatedDBObject(text, true);
        }
        acTrans.Commit();
      }
    }

    public static double GetPhaseConductorConstant(string wireSize)
    {
      switch (wireSize)
      {
        case "12":
          return 617;
        case "10":
          return 982;
        case "8":
          return 1559;
        case "6":
          return 2425;
        case "4":
          return 3806;
        case "3":
          return 4633;
        case "2":
          return 5907;
        case "1":
          return 7293;
        case "1/0":
          return 8925;
        case "2/0":
          return 10755;
        case "3/0":
          return 12844;
        case "4/0":
          return 15082;
        case "250 KCMIL":
          return 16483;
        case "300 KCMIL":
          return 18177;
        case "350 KCMIL":
          return 19704;
        case "400 KCMIL":
          return 20566;
        case "500 KCMIL":
          return 22185;
      }
      return 0;
    }

    public static double GetAicRatingFromTransformer(
      double kva,
      double powerFactor,
      double zValue,
      int length,
      int numPhaseConductors,
      double lineVoltage,
      string wireSize,
      bool is3Phase
    )
    {
      double yFactor = 2;
      if (is3Phase)
      {
        yFactor = 1.732;
      }
      double phaseConductorConstant = GetPhaseConductorConstant(wireSize);
      if (phaseConductorConstant == 0)
      {
        return 0;
      }
      double i = kva * 1000 / lineVoltage / (is3Phase ? yFactor : 1);
      double iSca = i * powerFactor / zValue;
      double fFactor =
        yFactor
        * Convert.ToDouble(length)
        * iSca
        / Convert.ToDouble(numPhaseConductors)
        / phaseConductorConstant
        / lineVoltage;
      double multiplier = 1.0 / (1.0 + fFactor);
      return iSca * multiplier;
    }

    public static double GetAicRating(
      double utilityFaultCurrent,
      int length,
      int numPhaseConductors,
      double lineVoltage,
      string wireSize,
      bool is3Phase
    )
    {
      double yFactor = 2;
      if (is3Phase)
      {
        yFactor = 1.732;
      }
      double phaseConductorConstant = GetPhaseConductorConstant(wireSize);
      if (phaseConductorConstant == 0)
      {
        return 0;
      }
      double fFactor =
        yFactor
        * Convert.ToDouble(length)
        * utilityFaultCurrent
        / Convert.ToDouble(numPhaseConductors)
        / phaseConductorConstant
        / lineVoltage;
      double multiplier = 1.0 / (1.0 + fFactor);

      return utilityFaultCurrent * multiplier;
    }

    [CommandMethod("HCND")]
    public void HCND()
    {
      HorizontalConduit = true;
      CND();
    }

    [CommandMethod("VCND")]
    public void VCND()
    {
      HorizontalConduit = false;
      CND();
    }

    [CommandMethod("CND")]
    public void CND()
    {
      bool horizontal = HorizontalConduit;
      var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      var ed = doc.Editor;
      Document acDoc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database acCurDb = acDoc.Database;

      var loadAmperagePrompt = new PromptStringOptions("\nEnter the load amperage: ");
      double loadAmperage = 0;
      var loadAmperageResult = ed.GetString(loadAmperagePrompt);
      while (loadAmperage <= 0)
      {
        if (loadAmperageResult.Status == PromptStatus.OK)
        {
          string loadAmperageString = loadAmperageResult.StringResult;
          if (double.TryParse(loadAmperageString, out double l))
          {
            loadAmperage = l;
            ed.WriteMessage($"\nLoad amperage set to {loadAmperage}");
          }
          else
          {
            ed.WriteMessage($"\nInvalid load amperage.");
            loadAmperageResult = ed.GetString(loadAmperagePrompt);
          }
        }
        else
        {
          return;
        }
      }

      var mocpPrompt = new PromptStringOptions("\nEnter the MOCP: ");
      var mocpResult = ed.GetString(mocpPrompt);
      int mocp = 0;
      while (mocp <= 0)
      {
        if (mocpResult.Status == PromptStatus.OK)
        {
          string breakerSizeString = mocpResult.StringResult;
          if (int.TryParse(breakerSizeString, out int s))
          {
            mocp = s;
            ed.WriteMessage($"\nMOCP to {mocp}");
          }
          else
          {
            ed.WriteMessage($"\nInvalid MOCP.");
            mocpResult = ed.GetString(mocpPrompt);
          }
        }
        else
        {
          return;
        }
      }

      if (Voltage <= 0)
      {
        SetVoltage();
        if (Voltage <= 0)
        {
          return;
        }
      }

      if (MaxVoltageDropPercent <= 0)
      {
        SetMaxVoltageDropPercent();
        if (Voltage <= 0)
        {
          return;
        }
      }

      if (Phase <= 0)
      {
        SetPhase();
        if (Phase <= 0)
        {
          return;
        }
      }
      var feederLengthPrompt = new PromptStringOptions("\nEnter the feeder length in feet: ");
      double distance = 0;
      var feederLengthResult = ed.GetString(feederLengthPrompt);
      while (distance <= 0)
      {
        if (feederLengthResult.Status == PromptStatus.OK)
        {
          string feederLengthString = feederLengthResult.StringResult;
          if (int.TryParse(feederLengthString, out int l))
          {
            distance = l;
            ed.WriteMessage($"\nDistance set to {distance}");
          }
          else
          {
            ed.WriteMessage($"\nInvalid distance.");
            feederLengthResult = ed.GetString(feederLengthPrompt);
          }
        }
        else
        {
          return;
        }
      }
      (
        string firstLine,
        string secondLine,
        string thirdLine,
        string supplemental1,
        string supplemental2,
        string supplemental3
      ) = GetWireAndConduitSizeText(
        loadAmperage,
        mocp,
        distance,
        Voltage,
        MaxVoltageDropPercent,
        Phase
      );
      // Prompt for a point
      PromptPointOptions ppo = new PromptPointOptions("\nSelect start point:");
      PromptPointResult ppr = ed.GetPoint(ppo);
      if (ppr.Status == PromptStatus.OK)
      {
        Point3d p = new Point3d(ppr.Value.X, ppr.Value.Y, ppr.Value.Z);
        AddWireAndConduitTextToPlan(
          acCurDb,
          p,
          firstLine,
          secondLine,
          thirdLine,
          supplemental1,
          supplemental2,
          supplemental3,
          horizontal
        );
      }
    }

    [CommandMethod("HR")]
    public void HR()
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Editor ed = doc.Editor;

      // Run the "AR" command
      ed.Command("AR");

      // Run the "SPOON" command
      ed.Command("SP");
    }

    [CommandMethod("CreateAttributeDef")]
    public void CreateAttributeDefinition()
    {
      // Get the current document and database
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      // Prompt for a point
      PromptPointResult ppr = ed.GetPoint("\nSpecify the insertion point for the attribute: ");
      if (ppr.Status != PromptStatus.OK)
        return;

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        // Open the BlockTable for read
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

        // Open the BlockTableRecord for write
        BlockTableRecord btr =
          tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

        // Create a new AttributeDefinition
        AttributeDefinition attDef = new AttributeDefinition
        {
          Position = ppr.Value,
          Height = 1.0,
          TextString = "Default",
          Tag = "TAG",
          Prompt = "Enter value for TAG: ",
          Justify = AttachmentPoint.BaseLeft, // Adjust alignment as necessary
        };

        // Add the attribute definition to the block table record
        btr.AppendEntity(attDef);
        tr.AddNewlyCreatedDBObject(attDef, true);

        // Commit the transaction
        tr.Commit();
      }
    }

    private void CreateMText(Point3d startPoint, AttachmentPoint attachmentPoint)
    {
      if (CADObjectCommands.Scale <= 0)
      {
        CADObjectCommands.SetScale();
        if (CADObjectCommands.Scale <= 0)
          return;
      }

      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Editor ed = doc.Editor;

      using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
      {
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite);

        double textHeight = 1.125 / CADObjectCommands.Scale;
        MText mText = new MText();
        mText.SetDatabaseDefaults();
        mText.TextHeight = textHeight;

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
        mText.Attachment = attachmentPoint;
        mText.Location = new Point3d(
          startPoint.X
            + (
              attachmentPoint == AttachmentPoint.TopLeft
                ? 0.25 / CADObjectCommands.Scale
                : (
                  attachmentPoint == AttachmentPoint.TopRight ? -0.25 / CADObjectCommands.Scale : 0
                )
            ),
          startPoint.Y + textHeight / 2,
          startPoint.Z
        );

        btr.AppendEntity(mText);
        tr.AddNewlyCreatedDBObject(mText, true);

        tr.Commit();

        doc.Editor.Command("_.MTEDIT", mText.ObjectId);
      }
    }

    private static void CreateTextWithJig(
      string layerName,
      TextHorizontalMode horizontalMode,
      string defaultText = null
    )
    {
      Document acDoc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database acCurDb = acDoc.Database;
      Editor ed = acDoc.Editor;

      if (Scale == -1.0)
      {
        SetScale();
      }

      if (Scale == -1.0)
      {
        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(
          "Please set the scale using the SetScale command before creating objects."
        );
        return;
      }

      double baseScale = 1.0 / 4.0;
      double baseTextHeight = 4.5;
      double textHeight = (baseScale / Scale) * baseTextHeight;

      string userText = defaultText;

      if (string.IsNullOrEmpty(userText))
      {
        PromptStringOptions promptStringOptions = new PromptStringOptions("\nEnter the text: ");
        PromptResult promptResult = ed.GetString(promptStringOptions);

        if (promptResult.Status != PromptStatus.OK)
        {
          ed.WriteMessage("\nText input canceled.");
          return;
        }

        userText = promptResult.StringResult;
      }

      GeneralTextJig jig = new GeneralTextJig(userText, textHeight, horizontalMode);
      PromptResult pr = ed.Drag(jig);

      if (pr.Status == PromptStatus.OK)
      {
        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        {
          BlockTable acBlkTbl =
            acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
          BlockTableRecord acBlkTblRec =
            acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite)
            as BlockTableRecord;

          DBText dbText = new DBText
          {
            Position = jig.InsertionPoint,
            TextString = userText,
            Height = textHeight,
            HorizontalMode = horizontalMode,
            Layer = layerName,
          };

          if (horizontalMode != TextHorizontalMode.TextLeft)
          {
            dbText.AlignmentPoint = jig.InsertionPoint;
          }

          TextStyleTable tst =
            acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
          if (tst.Has("rpm"))
          {
            dbText.TextStyleId = tst["rpm"];
          }
          else
          {
            ed.WriteMessage("\nText style 'rpm' not found.");
          }

          acBlkTblRec.AppendEntity(dbText);
          acTrans.AddNewlyCreatedDBObject(dbText, true);
          acTrans.Commit();
        }

        ed.WriteMessage(
          $"\nText '{userText}' created at {jig.InsertionPoint} with height {textHeight}."
        );
      }
      else
      {
        ed.WriteMessage("\nPoint selection canceled.");
      }
    }

    public static void JsonBlockCreator(string blockName)
    {
      if (Scale == -1.0)
      {
        SetScale();
      }

      if (Scale == -1.0)
      {
        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(
          "Please set the scale using the SetScale command before creating objects."
        );
        return;
      }

      if (PanelLocation == new Point3d(0, 0, 0) && blockName == "ar")
      {
        SetPanelLocation();
      }

      Document acDoc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database acCurDb = acDoc.Database;
      Editor ed = acDoc.Editor;

      string assemblyDirectory = System.IO.Path.GetDirectoryName(
        System.Reflection.Assembly.GetExecutingAssembly().Location
      );

      string jsonFilePath = System.IO.Path.Combine(
        assemblyDirectory,
        "Symbols",
        "BlockData",
        $"{blockName}.json"
      );

      if (File.Exists(jsonFilePath))
      {
        string jsonData = File.ReadAllText(jsonFilePath);

        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        {
          // Open the Block table for read
          BlockTable acBlkTbl =
            acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

          BlockTableRecord acBlkTblRec;

          if (acBlkTbl.Has(blockName + Scale.ToString()))
          {
            acBlkTblRec =
              acTrans.GetObject(acBlkTbl[blockName + Scale.ToString()], OpenMode.ForWrite)
              as BlockTableRecord;

            // Remove all entities from the block
            foreach (ObjectId id in acBlkTblRec)
            {
              acTrans.GetObject(id, OpenMode.ForWrite).Erase();
            }
          }
          else
          {
            acBlkTblRec = new BlockTableRecord { Name = blockName + Scale.ToString() };

            acBlkTbl.UpgradeOpen();
            acBlkTbl.Add(acBlkTblRec);
            acTrans.AddNewlyCreatedDBObject(acBlkTblRec, true);
          }

          // Set the block table record as the current space
          acCurDb.TransactionManager.TopTransaction.GetObject(
            SymbolUtilityServices.GetBlockModelSpaceId(acCurDb),
            OpenMode.ForWrite
          );

          CreateObjectFromData(jsonData, Point3d.Origin, acBlkTblRec);

          acTrans.Commit();
        }

        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        {
          BlockTable acBlkTbl =
            acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

          var acBlkTblRec =
            acTrans.GetObject(acBlkTbl[blockName + Scale.ToString()], OpenMode.ForWrite)
            as BlockTableRecord;
          PromptPointResult ppr = ed.GetPoint("\nSpecify insertion point: ");

          if (ppr.Status == PromptStatus.OK)
          {
            Point3d insertionPoint = ppr.Value;

            ObjectData objectData = JsonConvert.DeserializeObject<ObjectData>(jsonData);

            using (
              BlockReference acBlkRef = new BlockReference(insertionPoint, acBlkTblRec.ObjectId)
            )
            {
              BlockTableRecord currentSpace =
                acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
              currentSpace.AppendEntity(acBlkRef);
              acTrans.AddNewlyCreatedDBObject(acBlkRef, true);

              double savedScale = objectData.Scale;
              double scaleFactor = CalculateScaleFactor(savedScale, Scale);

              foreach (var attDef in objectData.AttributeDefinitions)
              {
                var attDefObj = CreateAttributeDefinition(
                  Point3d.Origin,
                  acTrans,
                  acBlkTblRec,
                  attDef,
                  scaleFactor
                );
                CreateAttributeReference(acBlkRef, attDef, attDefObj, acTrans, scaleFactor);
              }
            }

            acTrans.Commit();
            ed.WriteMessage($"\nBlock '{blockName}' created at {insertionPoint}.");
          }
          else
          {
            ed.WriteMessage("\nBlock creation canceled.");
          }
        }
      }
      else
      {
        ed.WriteMessage($"\nBlock '{blockName}' not found in BlockData directory.");
      }
    }

    public static void JsonObjectCreator(string blockName)
    {
      if (Scale == -1.0)
      {
        SetScale();
      }

      if (Scale == -1.0)
      {
        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(
          "Please set the scale using the SetScale command before creating objects."
        );
        return;
      }

      if (PanelLocation == new Point3d(0, 0, 0) && blockName == "ar")
      {
        SetPanelLocation();
      }

      Document acDoc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database acCurDb = acDoc.Database;
      Editor ed = acDoc.Editor;

      ed.WriteMessage($"\nCurrent panel location is set to {PanelLocation}");
      ed.WriteMessage("\nRun the 'SetPanelLocation' command to set a new panel location.");

      string assemblyDirectory = System.IO.Path.GetDirectoryName(
        System.Reflection.Assembly.GetExecutingAssembly().Location
      );
      string jsonFilePath = System.IO.Path.Combine(
        assemblyDirectory,
        "Symbols",
        "BlockData",
        $"{blockName}.json"
      );

      if (File.Exists(jsonFilePath))
      {
        string jsonData = File.ReadAllText(jsonFilePath);

        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
        {
          PromptPointResult ppr = ed.GetPoint("\nSpecify insertion point: ");

          if (ppr.Status == PromptStatus.OK)
          {
            Point3d insertionPoint = ppr.Value;

            BlockTableRecord currentSpace =
              acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

            // Insert the object at the specified insertion point
            CreateObjectFromData(jsonData, insertionPoint, currentSpace);

            acTrans.Commit();

            ed.WriteMessage($"\nObject '{blockName}' created at {insertionPoint}.");
          }
          else
          {
            ed.WriteMessage("\nObject creation canceled.");
          }
        }
      }
      else
      {
        ed.WriteMessage($"\nObject data for '{blockName}' not found in BlockData directory.");
      }
    }

    public static double CalculateScaleFactor(double savedScale, double currentScale)
    {
      return savedScale / currentScale;
    }

    public static void CreateAttributeReference(
      BlockReference blockRef,
      AttributeDefinitionData attDefData,
      AttributeDefinition attDefObj,
      Transaction acTrans,
      double scaleFactor
    )
    {
      AttributeReference attRef = new AttributeReference();
      attRef.SetAttributeFromBlock(attDefObj, blockRef.BlockTransform);
      attRef.TextString = attDefData.Tag;
      attRef.Tag = attDefData.Tag;
      attRef.Position = new Point3d(
        attDefData.Location.X * scaleFactor,
        attDefData.Location.Y * scaleFactor,
        attDefData.Location.Z * scaleFactor
      );
      attRef.Height = attDefData.Height * scaleFactor;
      attRef.Rotation = attDefData.Rotation;
      blockRef.AttributeCollection.AppendAttribute(attRef);
      acTrans.AddNewlyCreatedDBObject(attRef, true);
    }

    public static HatchData HatchBoundary(Hatch hatch)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Editor ed = doc.Editor;

      var hatchData = new HatchData
      {
        Layer = hatch.Layer,
        PatternName = hatch.PatternName,
        PatternScale = hatch.PatternScale,
        PatternType = hatch.PatternType,
        Angle = hatch.PatternAngle,
        Polylines = new List<PolylineData>(),
        Lines = new List<LineData>(),
        Arcs = new List<ArcData>(),
        Circles = new List<CircleData>(),
        Ellipses = new List<EllipseData>(),
        Splines = new List<SplineData>(),
      };

      var plane = hatch.GetPlane();
      var normal = plane.Normal;
      var nLoops = hatch.NumberOfLoops;

      ed.WriteMessage($"\nNumber of loops: {nLoops}");

      using (Transaction tr = doc.TransactionManager.StartTransaction())
      {
        if (hatch != null)
        {
          BlockTableRecord btr = tr.GetObject(hatch.OwnerId, OpenMode.ForWrite) as BlockTableRecord;
          if (btr == null)
            return hatchData;

          for (int i = 0; i < nLoops; i++)
          {
            HatchLoop loop = hatch.GetLoopAt(i);
            ed.WriteMessage($"\nProcessing loop {i + 1}/{nLoops}");

            if (loop.IsPolyline)
            {
              ed.WriteMessage("\nLoop is a polyline.");
              var polylineData = new PolylineData
              {
                Layer = hatch.Layer,
                Vectors = new List<SimpleVector3d>(),
                LineType = "SOLID",
                Closed = false,
              };

              foreach (BulgeVertex bv in loop.Polyline)
              {
                polylineData.Vectors.Add(new SimpleVector3d(bv.Vertex.X, bv.Vertex.Y, 0));
              }
              hatchData.Polylines.Add(polylineData);
            }
            else
            {
              ed.WriteMessage("\nLoop is not a polyline.");
              foreach (Curve2d cv in loop.Curves)
              {
                LineSegment2d line2d = cv as LineSegment2d;
                CircularArc2d arc2d = cv as CircularArc2d;
                EllipticalArc2d ellipse2d = cv as EllipticalArc2d;
                NurbCurve2d spline2d = cv as NurbCurve2d;

                if (line2d != null)
                {
                  ed.WriteMessage("\nFound a line segment.");
                  var lineData = new LineData
                  {
                    Layer = hatch.Layer,
                    StartPoint = new SimpleVector3d(line2d.StartPoint.X, line2d.StartPoint.Y, 0),
                    EndPoint = new SimpleVector3d(line2d.EndPoint.X, line2d.EndPoint.Y, 0),
                  };
                  hatchData.Lines.Add(lineData);
                }
                else if (arc2d != null)
                {
                  ed.WriteMessage("\nFound a circular arc.");
                  if (Math.Abs(arc2d.StartAngle - arc2d.EndAngle) < 1e-5)
                  {
                    var circleData = new CircleData
                    {
                      Layer = hatch.Layer,
                      Center = new SimpleVector3d(arc2d.Center.X, arc2d.Center.Y, 0),
                      Radius = arc2d.Radius,
                    };
                    hatchData.Circles.Add(circleData);
                  }
                  else
                  {
                    double angle = new Vector3d(plane, arc2d.ReferenceVector).AngleOnPlane(plane);
                    var arcData = new ArcData
                    {
                      Layer = hatch.Layer,
                      Center = new SimpleVector3d(arc2d.Center.X, arc2d.Center.Y, 0),
                      Radius = arc2d.Radius,
                      StartAngle = arc2d.StartAngle + angle,
                      EndAngle = arc2d.EndAngle + angle,
                    };
                    hatchData.Arcs.Add(arcData);
                  }
                }
                else if (ellipse2d != null)
                {
                  ed.WriteMessage("\nFound an elliptical arc.");
                  var ellipseData = new EllipseData
                  {
                    Layer = hatch.Layer,
                    UnitNormal = new SimpleVector3d(normal.X, normal.Y, normal.Z),
                    Center = new SimpleVector3d(ellipse2d.Center.X, ellipse2d.Center.Y, 0),
                    MajorAxis = new SimpleVector3d(ellipse2d.MajorAxis.X, ellipse2d.MajorAxis.Y, 0),
                    MajorRadius = ellipse2d.MajorRadius,
                    MinorRadius = ellipse2d.MinorRadius,
                    StartAngle = ellipse2d.StartAngle,
                    EndAngle = ellipse2d.EndAngle,
                  };
                  hatchData.Ellipses.Add(ellipseData);
                }
                else if (spline2d != null)
                {
                  ed.WriteMessage("\nFound a spline.");
                  var splineData = new SplineData { Layer = hatch.Layer };

                  if (spline2d.HasFitData)
                  {
                    NurbCurve2dFitData n2fd = spline2d.FitData;
                    splineData.FitPoints = new List<SimpleVector3d>();
                    foreach (Point2d p in n2fd.FitPoints)
                    {
                      splineData.FitPoints.Add(new SimpleVector3d(p.X, p.Y, 0));
                    }
                    splineData.StartTangent = new SimpleVector3d(
                      n2fd.StartTangent.X,
                      n2fd.StartTangent.Y,
                      0
                    );
                    splineData.EndTangent = new SimpleVector3d(
                      n2fd.EndTangent.X,
                      n2fd.EndTangent.Y,
                      0
                    );
                    splineData.Degree = n2fd.Degree;
                    splineData.FitTolerance = n2fd.FitTolerance.EqualPoint;
                  }
                  else
                  {
                    NurbCurve2dData n2fd = spline2d.DefinitionData;
                    splineData.ControlPoints = new List<SimpleVector3d>();
                    foreach (Point2d p in n2fd.ControlPoints)
                    {
                      splineData.ControlPoints.Add(new SimpleVector3d(p.X, p.Y, 0));
                    }
                    splineData.Knots = new List<double>(n2fd.Knots.Count);
                    foreach (double k in n2fd.Knots)
                    {
                      splineData.Knots.Add(k);
                    }
                    splineData.Degree = n2fd.Degree;
                    splineData.Rational = n2fd.Rational;
                    splineData.Periodic = spline2d.IsPeriodic(out double period);
                    splineData.Closed = spline2d.IsClosed();
                    splineData.Weights = new List<double>();
                    splineData.Weights.AddRange((IEnumerable<double>)n2fd.Weights);
                    splineData.KnotTolerance = n2fd.Knots.Tolerance;
                  }

                  hatchData.Splines.Add(splineData);
                }
              }
            }
          }
        }

        for (int i = 0; i < hatchData.Arcs.Count; i++)
        {
          var arc = hatchData.Arcs[i];

          if (Math.Abs(arc.StartAngle) < 1e-5 && Math.Abs(arc.EndAngle - (2 * Math.PI)) < 1e-5)
          {
            var circleData = new CircleData
            {
              Layer = hatch.Layer,
              Center = arc.Center,
              Radius = arc.Radius,
            };
            hatchData.Circles.Add(circleData);
            hatchData.Arcs.RemoveAt(i);
            i--;
          }
        }

        tr.Commit();

        return hatchData;
      }
    }

    public static void SetTextPositionAndAlignment(
      Point3d basePoint,
      double scaleFactor,
      TextData textData,
      dynamic textObject
    )
    {
      if (textObject is DBText dbText)
      {
        if (
          textData.HorizontalMode == TextHorizontalMode.TextLeft
          || textData.HorizontalMode == TextHorizontalMode.TextMid
        )
        {
          dbText.Position = new Point3d(
            (basePoint.X + textData.Location.X) * scaleFactor,
            (basePoint.Y + textData.Location.Y) * scaleFactor,
            (basePoint.Z + textData.Location.Z) * scaleFactor
          );
        }
        else if (textData.HorizontalMode == TextHorizontalMode.TextCenter)
        {
          dbText.HorizontalMode = textData.HorizontalMode;
          dbText.AlignmentPoint = new Point3d(
            (basePoint.X + textData.AlignmentPoint.X) * scaleFactor,
            (basePoint.Y + textData.AlignmentPoint.Y) * scaleFactor,
            (basePoint.Z + textData.AlignmentPoint.Z) * scaleFactor
          );
          dbText.Justify = textData.Justification;
        }
        else
        {
          dbText.HorizontalMode = textData.HorizontalMode;
          dbText.AlignmentPoint = new Point3d(
            (basePoint.X + textData.AlignmentPoint.X) * scaleFactor,
            (basePoint.Y + textData.AlignmentPoint.Y) * scaleFactor,
            (basePoint.Z + textData.AlignmentPoint.Z) * scaleFactor
          );
        }
      }
      else if (textObject is AttributeDefinition attDef)
      {
        if (
          textData.HorizontalMode == TextHorizontalMode.TextLeft
          || textData.HorizontalMode == TextHorizontalMode.TextMid
        )
        {
          attDef.Position = new Point3d(
            (basePoint.X + textData.Location.X) * scaleFactor,
            (basePoint.Y + textData.Location.Y) * scaleFactor,
            (basePoint.Z + textData.Location.Z) * scaleFactor
          );
        }
        else if (textData.HorizontalMode == TextHorizontalMode.TextCenter)
        {
          attDef.HorizontalMode = textData.HorizontalMode;
          attDef.AlignmentPoint = new Point3d(
            (basePoint.X + textData.AlignmentPoint.X) * scaleFactor,
            (basePoint.Y + textData.AlignmentPoint.Y) * scaleFactor,
            (basePoint.Z + textData.AlignmentPoint.Z) * scaleFactor
          );
          attDef.Justify = textData.Justification;
        }
        else
        {
          attDef.HorizontalMode = textData.HorizontalMode;
          attDef.AlignmentPoint = new Point3d(
            (basePoint.X + textData.AlignmentPoint.X) * scaleFactor,
            (basePoint.Y + textData.AlignmentPoint.Y) * scaleFactor,
            (basePoint.Z + textData.AlignmentPoint.Z) * scaleFactor
          );
        }
      }
    }

    public static AttributeDefinition CreateAttributeDefinition(
      Point3d basePoint,
      Transaction acTrans,
      BlockTableRecord acBlkTblRec,
      AttributeDefinitionData attDefData,
      double scaleFactor
    )
    {
      string tag = attDefData.Tag;

      // Check if an attribute definition with the same tag already exists in the block
      foreach (ObjectId objId in acBlkTblRec)
      {
        DBObject obj = acTrans.GetObject(objId, OpenMode.ForRead);
        if (obj is AttributeDefinition existingAttDef && existingAttDef.Tag == tag)
        {
          // Update the existing attribute definition
          existingAttDef.UpgradeOpen();
          existingAttDef.Invisible = attDefData.Invisible;
          existingAttDef.Prompt = attDefData.Prompt;
          existingAttDef.TextString = attDefData.Contents;
          existingAttDef.Height = attDefData.Height * scaleFactor;
          existingAttDef.Rotation = attDefData.Rotation;
          SetTextStyleByName(existingAttDef, attDefData.Style);

          SetTextPositionAndAlignment(basePoint, scaleFactor, attDefData, existingAttDef);

          return existingAttDef;
        }
      }

      // If no existing attribute definition is found, create a new one
      AttributeDefinition attDef = new AttributeDefinition();
      var textStyleObject = new TextStyle(attDefData.StyleAttributes);
      textStyleObject.CreateStyleIfNotExisting(attDefData.Style);

      attDef.Layer = attDefData.Layer;
      attDef.TextString = attDefData.Contents;
      attDef.Height = attDefData.Height * scaleFactor;
      attDef.Rotation = attDefData.Rotation;
      attDef.Tag = attDefData.Tag;
      attDef.Prompt = attDefData.Prompt;
      attDef.Invisible = false;

      SetTextStyleByName(attDef, attDefData.Style);

      SetTextPositionAndAlignment(basePoint, scaleFactor, attDefData, attDef);

      acBlkTblRec.AppendEntity(attDef);
      acTrans.AddNewlyCreatedDBObject(attDef, true);

      return attDef;
    }

    public static Point3d CreateArc(
      Point3d basePoint,
      Transaction acTrans,
      BlockTableRecord acBlkTblRec,
      ArcData arcData,
      double scaleFactor
    )
    {
      Arc arc = new Arc();
      arc.Layer = arcData.Layer;
      arc.Center = new Point3d(
        basePoint.X + arcData.Center.X * scaleFactor,
        basePoint.Y + arcData.Center.Y * scaleFactor,
        basePoint.Z + arcData.Center.Z * scaleFactor
      );
      arc.Radius = arcData.Radius * scaleFactor;
      arc.StartAngle = arcData.StartAngle;
      arc.EndAngle = arcData.EndAngle;

      acBlkTblRec.AppendEntity(arc);
      acTrans.AddNewlyCreatedDBObject(arc, true);
      return basePoint;
    }

    public static Point3d CreateCircle(
      Point3d basePoint,
      Transaction acTrans,
      BlockTableRecord acBlkTblRec,
      CircleData circleData,
      double scaleFactor
    )
    {
      Circle circle = new Circle();
      circle.Layer = circleData.Layer;
      circle.Center = new Point3d(
        basePoint.X + circleData.Center.X * scaleFactor,
        basePoint.Y + circleData.Center.Y * scaleFactor,
        basePoint.Z + circleData.Center.Z * scaleFactor
      );
      circle.Radius = circleData.Radius * scaleFactor;

      acBlkTblRec.AppendEntity(circle);
      acTrans.AddNewlyCreatedDBObject(circle, true);
      return basePoint;
    }

    public static Point3d CreateEllipse(
      Point3d basePoint,
      Transaction acTrans,
      BlockTableRecord acBlkTblRec,
      EllipseData ellipseData,
      double scaleFactor
    )
    {
      Ellipse ellipse = new Ellipse();
      ellipse.Layer = ellipseData.Layer;
      Point3d center = new Point3d(
        basePoint.X + ellipseData.Center.X * scaleFactor,
        basePoint.Y + ellipseData.Center.Y * scaleFactor,
        basePoint.Z + ellipseData.Center.Z * scaleFactor
      );
      Vector3d majorAxis = new Vector3d(
        ellipseData.MajorAxis.X * scaleFactor,
        ellipseData.MajorAxis.Y * scaleFactor,
        ellipseData.MajorAxis.Z * scaleFactor
      );
      double radiusRatio = ellipseData.RadiusRatio();
      double startAngle = ellipseData.StartAngle;
      double endAngle = ellipseData.EndAngle;
      Vector3d unitNormal = new Vector3d(0, 0, 1);

      ellipse.Set(center, unitNormal, majorAxis, radiusRatio, startAngle, endAngle);

      acBlkTblRec.AppendEntity(ellipse);
      acTrans.AddNewlyCreatedDBObject(ellipse, true);
      return basePoint;
    }

    public static Point3d CreateLine(
      Point3d basePoint,
      Transaction acTrans,
      BlockTableRecord acBlkTblRec,
      LineData lineData,
      double scaleFactor,
      string lineType = ""
    )
    {
      Line line = new Line();
      line.Layer = lineData.Layer;
      line.StartPoint = new Point3d(
        basePoint.X + lineData.StartPoint.X * scaleFactor,
        basePoint.Y + lineData.StartPoint.Y * scaleFactor,
        basePoint.Z + lineData.StartPoint.Z * scaleFactor
      );
      line.EndPoint = new Point3d(
        basePoint.X + lineData.EndPoint.X * scaleFactor,
        basePoint.Y + lineData.EndPoint.Y * scaleFactor,
        basePoint.Z + lineData.EndPoint.Z * scaleFactor
      );
      if (!String.IsNullOrEmpty(lineType))
      {
        line.Linetype = lineType;
      }

      acBlkTblRec.AppendEntity(line);
      acTrans.AddNewlyCreatedDBObject(line, true);
      return basePoint;
    }

    public static Point3d CreateMText(
      Point3d basePoint,
      Transaction acTrans,
      BlockTableRecord acBlkTblRec,
      MTextData mTextData,
      double scaleFactor
    )
    {
      MText mText = new MText();
      mText.Layer = mTextData.Layer;

      var textStyleObject = new TextStyle(mTextData.StyleAttributes);
      textStyleObject.CreateStyleIfNotExisting(mTextData.Style);
      SetTextStyleByName(mText, mTextData.Style);

      mText.Attachment = (AttachmentPoint)
        Enum.Parse(typeof(AttachmentPoint), mTextData.Justification);
      mText.Contents = mTextData.Contents;
      mText.Location = new Point3d(
        basePoint.X + mTextData.Location.X * scaleFactor,
        basePoint.Y + mTextData.Location.Y * scaleFactor,
        basePoint.Z + mTextData.Location.Z * scaleFactor
      );
      mText.TextHeight = mTextData.TextHeight * scaleFactor;
      mText.Width = mTextData.Width * scaleFactor;

      acBlkTblRec.AppendEntity(mText);
      acTrans.AddNewlyCreatedDBObject(mText, true);
      return basePoint;
    }

    public static Point3d CreatePolyline(
      Point3d basePoint,
      Transaction acTrans,
      BlockTableRecord acBlkTblRec,
      PolylineData polylineData,
      double scaleFactor
    )
    {
      Polyline polyline = new Polyline();
      polyline.Layer = polylineData.Layer;
      polyline.Linetype = polylineData.LineType;
      polyline.Closed = polylineData.Closed;

      for (int i = 0; i < polylineData.Vectors.Count; i++)
      {
        SimpleVector3d vector = polylineData.Vectors[i];
        double bulge = polylineData.Bulges[i];
        polyline.AddVertexAt(
          i,
          new Point2d(basePoint.X + vector.X * scaleFactor, basePoint.Y + vector.Y * scaleFactor),
          bulge,
          polylineData.StartWidth * scaleFactor,
          polylineData.EndWidth * scaleFactor
        );
      }

      polyline.ConstantWidth = polylineData.GlobalWidth * scaleFactor;

      acBlkTblRec.AppendEntity(polyline);
      acTrans.AddNewlyCreatedDBObject(polyline, true);
      return basePoint;
    }

    public static Point3d CreatePolyline2d(
      Point3d basePoint,
      Transaction acTrans,
      BlockTableRecord acBlkTblRec,
      Polyline2dData polylineData,
      double scaleFactor
    )
    {
      // Create a Point3dCollection
      Point3dCollection pointCollection = new Point3dCollection();

      // Create and add points to the Point3dCollection
      for (int i = 0; i < polylineData.Vertices.Count; i++)
      {
        var vector = polylineData.Vertices[i];
        Point3d point = new Point3d(
          basePoint.X + vector.X * scaleFactor,
          basePoint.Y + vector.Y * scaleFactor,
          basePoint.Z + vector.Z * scaleFactor
        );
        pointCollection.Add(point);
      }

      // Create a new Polyline2d using the Point3dCollection
      Polyline2d polyline = new Polyline2d(
        Poly2dType.SimplePoly,
        pointCollection,
        0,
        false,
        0,
        0,
        null
      );
      polyline.Layer = polylineData.Layer;

      // Append the Polyline2d to the BlockTableRecord
      acBlkTblRec.AppendEntity(polyline);
      acTrans.AddNewlyCreatedDBObject(polyline, true);

      return basePoint;
    }

    public static Point3d CreateSolid(
      Point3d basePoint,
      Transaction acTrans,
      BlockTableRecord acBlkTblRec,
      SolidData solidData,
      double scaleFactor
    )
    {
      Solid solid = new Solid();
      solid.Layer = solidData.Layer;
      for (short i = 0; i < solidData.Vertices.Count; i++)
      {
        SimpleVector3d vector = solidData.Vertices[i];
        solid.SetPointAt(
          i,
          new Point3d(
            basePoint.X + vector.X * scaleFactor,
            basePoint.Y + vector.Y * scaleFactor,
            basePoint.Z + vector.Z * scaleFactor
          )
        );
      }

      acBlkTblRec.AppendEntity(solid);
      acTrans.AddNewlyCreatedDBObject(solid, true);
      return basePoint;
    }

    public static Point3d CreateText(
      Point3d basePoint,
      Transaction acTrans,
      BlockTableRecord acBlkTblRec,
      TextData text,
      double scaleFactor
    )
    {
      Autodesk.AutoCAD.EditorInput.Editor ed = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument
        .Editor;

      DBText dbText = new DBText();
      var textStyleObject = new TextStyle(text.StyleAttributes);
      textStyleObject.CreateStyleIfNotExisting(text.Style);

      dbText.Layer = text.Layer;
      dbText.TextString = text.Contents;
      ed.WriteMessage($"This is the scalefactor inside createtext function: {scaleFactor}");
      dbText.Height = text.Height * scaleFactor;
      dbText.Rotation = text.Rotation;

      SetTextStyleByName(dbText, text.Style);

      SetTextPositionAndAlignment(basePoint, scaleFactor, text, dbText);

      acBlkTblRec.AppendEntity(dbText);
      acTrans.AddNewlyCreatedDBObject(dbText, true);

      return basePoint;
    }

    public static void CreateHatch(
      Point3d basePoint,
      Transaction acTrans,
      BlockTableRecord acBlkTblRec,
      HatchData hatchData,
      double scaleFactor
    )
    {
      Hatch hatch = new Hatch();
      hatch.SetDatabaseDefaults();
      hatch.Layer = hatchData.Layer;
      hatch.SetHatchPattern(HatchPatternType.PreDefined, hatchData.PatternName);
      hatch.HatchStyle = Autodesk.AutoCAD.DatabaseServices.HatchStyle.Ignore;
      if (hatchData.PatternName != "SOLID")
      {
        hatch.PatternScale = hatchData.PatternScale * scaleFactor;
        hatch.PatternAngle = hatchData.Angle;
      }

      ObjectIdCollection loopIds = new ObjectIdCollection();

      foreach (var polyline in hatchData.Polylines)
      {
        using (Polyline poly = new Polyline())
        {
          for (int i = 0; i < polyline.Vectors.Count; i++)
          {
            poly.AddVertexAt(
              i,
              new Point2d(
                basePoint.X + polyline.Vectors[i].X * scaleFactor,
                basePoint.Y + polyline.Vectors[i].Y * scaleFactor
              ),
              0,
              0,
              0
            );
          }
          poly.Closed = polyline.Closed;

          acBlkTblRec.AppendEntity(poly);
          acTrans.AddNewlyCreatedDBObject(poly, true);

          loopIds.Add(poly.ObjectId);
        }
      }

      foreach (var line in hatchData.Lines)
      {
        using (
          Line ln = new Line(
            new Point3d(
              basePoint.X + line.StartPoint.X * scaleFactor,
              basePoint.Y + line.StartPoint.Y * scaleFactor,
              basePoint.Z + line.StartPoint.Z * scaleFactor
            ),
            new Point3d(
              basePoint.X + line.EndPoint.X * scaleFactor,
              basePoint.Y + line.EndPoint.Y * scaleFactor,
              basePoint.Z + line.EndPoint.Z * scaleFactor
            )
          )
        )
        {
          acBlkTblRec.AppendEntity(ln);
          acTrans.AddNewlyCreatedDBObject(ln, true);

          loopIds.Add(ln.ObjectId);
        }
      }

      foreach (var arc in hatchData.Arcs)
      {
        using (
          Arc arc3d = new Arc(
            new Point3d(
              basePoint.X + arc.Center.X * scaleFactor,
              basePoint.Y + arc.Center.Y * scaleFactor,
              basePoint.Z + arc.Center.Z * scaleFactor
            ),
            arc.Radius * scaleFactor,
            arc.StartAngle,
            arc.EndAngle
          )
        )
        {
          acBlkTblRec.AppendEntity(arc3d);
          acTrans.AddNewlyCreatedDBObject(arc3d, true);

          loopIds.Add(arc3d.ObjectId);
        }
      }

      foreach (var circle in hatchData.Circles)
      {
        using (
          Circle circle3d = new Circle(
            new Point3d(
              basePoint.X + circle.Center.X * scaleFactor,
              basePoint.Y + circle.Center.Y * scaleFactor,
              basePoint.Z + circle.Center.Z * scaleFactor
            ),
            Vector3d.ZAxis,
            circle.Radius * scaleFactor
          )
        )
        {
          acBlkTblRec.AppendEntity(circle3d);
          acTrans.AddNewlyCreatedDBObject(circle3d, true);

          loopIds.Add(circle3d.ObjectId);
        }
      }

      hatch.AppendLoop(HatchLoopTypes.External, loopIds);
      hatch.EvaluateHatch(true);
      acBlkTblRec.AppendEntity(hatch);
      acTrans.AddNewlyCreatedDBObject(hatch, true);
    }

    private static void SetTextStyleByName(Entity textEntity, string styleName)
    {
      if (!(textEntity is MText || textEntity is DBText))
      {
        throw new ArgumentException("The textEntity must be of type MText or DBText.");
      }

      Database db = HostApplicationServices.WorkingDatabase;
      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        TextStyleTable textStyleTable =
          tr.GetObject(db.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
        if (textStyleTable.Has(styleName))
        {
          TextStyleTableRecord textStyle =
            tr.GetObject(textStyleTable[styleName], OpenMode.ForRead) as TextStyleTableRecord;
          if (textEntity is MText mTextEntity)
          {
            mTextEntity.TextStyleId = textStyle.ObjectId;
          }
          else if (textEntity is DBText dbTextEntity)
          {
            dbTextEntity.TextStyleId = textStyle.ObjectId;
          }
        }
        tr.Commit();
      }
    }

    public static void CreateObjectFromData(
      string jsonData,
      Point3d basePoint,
      BlockTableRecord block
    )
    {
      if (Scale == -1.0)
      {
        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(
          "Please set the scale using the SetScale command before creating objects."
        );
        return;
      }

      ObjectData objectData = JsonConvert.DeserializeObject<ObjectData>(jsonData);

      double savedScale = objectData.Scale;
      double scaleFactor = CalculateScaleFactor(savedScale, Scale);

      Document acDoc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database acCurDb = acDoc.Database;

      using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
      {
        BlockTableRecord acBlkTblRec = block;

        foreach (var polyline in objectData.Polylines)
        {
          basePoint = CreatePolyline(basePoint, acTrans, acBlkTblRec, polyline, scaleFactor);
        }

        foreach (var polyline2d in objectData.Polylines2d)
        {
          basePoint = CreatePolyline2d(basePoint, acTrans, acBlkTblRec, polyline2d, scaleFactor);
        }

        foreach (var line in objectData.Lines)
        {
          basePoint = CreateLine(basePoint, acTrans, acBlkTblRec, line, scaleFactor);
        }

        foreach (var arc in objectData.Arcs)
        {
          basePoint = CreateArc(basePoint, acTrans, acBlkTblRec, arc, scaleFactor);
        }

        foreach (var circle in objectData.Circles)
        {
          basePoint = CreateCircle(basePoint, acTrans, acBlkTblRec, circle, scaleFactor);
        }

        foreach (var ellipse in objectData.Ellipses)
        {
          basePoint = CreateEllipse(basePoint, acTrans, acBlkTblRec, ellipse, scaleFactor);
        }

        foreach (var mText in objectData.MTexts)
        {
          basePoint = CreateMText(basePoint, acTrans, acBlkTblRec, mText, scaleFactor);
        }

        foreach (var text in objectData.Texts)
        {
          basePoint = CreateText(basePoint, acTrans, acBlkTblRec, text, scaleFactor);
        }

        foreach (var solid in objectData.Solids)
        {
          basePoint = CreateSolid(basePoint, acTrans, acBlkTblRec, solid, scaleFactor);
        }

        foreach (var hatch in objectData.Hatches)
        {
          CreateHatch(basePoint, acTrans, acBlkTblRec, hatch, scaleFactor);
        }

        acTrans.Commit();
      }
    }

    public static ObjectData HandleAttributeDefinition(
      AttributeDefinition attDef,
      ObjectData data,
      Point3d origin
    )
    {
      var attDefData = new AttributeDefinitionData
      {
        Layer = attDef.Layer,
        Tag = attDef.Tag,
        Prompt = attDef.Prompt,
        Contents = attDef.TextString,
        Location = new SimpleVector3d
        {
          X = attDef.Position.X - origin.X,
          Y = attDef.Position.Y - origin.Y,
          Z = attDef.Position.Z - origin.Z,
        },
        Height = attDef.Height,
        Rotation = attDef.Rotation,
        Invisible = attDef.Invisible,
      };

      data.AttributeDefinitions.Add(attDefData);

      return data;
    }

    public static ObjectData HandleArc(Arc arc, ObjectData data, Point3d origin)
    {
      var arcData = new ArcData
      {
        Layer = arc.Layer,
        Center = new SimpleVector3d
        {
          X = arc.Center.X - origin.X,
          Y = arc.Center.Y - origin.Y,
          Z = arc.Center.Z - origin.Z,
        },
        Radius = arc.Radius,
        StartAngle = arc.StartAngle,
        EndAngle = arc.EndAngle,
      };

      data.Arcs.Add(arcData);

      return data;
    }

    public static ObjectData HandleCircle(Circle circle, ObjectData data, Point3d origin)
    {
      var circleData = new CircleData
      {
        Layer = circle.Layer,
        Center = new SimpleVector3d
        {
          X = circle.Center.X - origin.X,
          Y = circle.Center.Y - origin.Y,
          Z = circle.Center.Z - origin.Z,
        },
        Radius = circle.Radius,
      };

      data.Circles.Add(circleData);

      return data;
    }

    public static ObjectData HandleEllipse(Ellipse ellipse, ObjectData data, Point3d origin)
    {
      var ellipseData = new EllipseData
      {
        Layer = ellipse.Layer,
        UnitNormal = new SimpleVector3d
        {
          X = ellipse.Normal.X,
          Y = ellipse.Normal.Y,
          Z = ellipse.Normal.Z,
        },
        Center = new SimpleVector3d
        {
          X = ellipse.Center.X - origin.X,
          Y = ellipse.Center.Y - origin.Y,
          Z = ellipse.Center.Z - origin.Z,
        },
        MajorAxis = new SimpleVector3d
        {
          X = ellipse.MajorAxis.X,
          Y = ellipse.MajorAxis.Y,
          Z = ellipse.MajorAxis.Z,
        },
        MajorRadius = ellipse.MajorRadius,
        MinorRadius = ellipse.MinorRadius,
        StartAngle = ellipse.StartAngle,
        EndAngle = ellipse.EndAngle,
      };

      data.Ellipses.Add(ellipseData);

      return data;
    }

    public static ObjectData HandleLine(Line line, ObjectData data, Point3d origin)
    {
      var lineData = new LineData
      {
        Layer = line.Layer,
        StartPoint = new SimpleVector3d
        {
          X = line.StartPoint.X - origin.X,
          Y = line.StartPoint.Y - origin.Y,
          Z = line.StartPoint.Z - origin.Z,
        },
        EndPoint = new SimpleVector3d
        {
          X = line.EndPoint.X - origin.X,
          Y = line.EndPoint.Y - origin.Y,
          Z = line.EndPoint.Z - origin.Z,
        },
      };

      data.Lines.Add(lineData);

      return data;
    }

    public static ObjectData HandleMText(MText mText, ObjectData data, Point3d origin)
    {
      var mTextData = new MTextData
      {
        Layer = mText.Layer,
        Style = mText.TextStyleName,
        Justification = mText.Attachment.ToString(),
        Contents = mText.Contents,
        Location = new SimpleVector3d
        {
          X = mText.Location.X - origin.X,
          Y = mText.Location.Y - origin.Y,
          Z = mText.Location.Z - origin.Z,
        },
        LineSpaceDistance = mText.LineSpaceDistance,
        LineSpaceFactor = mText.LineSpacingFactor,
        LineSpacingStyle = mText.LineSpacingStyle,
        TextHeight = mText.TextHeight,
        Width = mText.Width,
        Rotation = mText.Rotation,
        Direction = mText.Direction,
        StyleAttributes = new StyleAttributes
        {
          Height = mText.TextHeight,
          FontName = mText.TextStyleName,
        },
      };

      data.MTexts.Add(mTextData);

      return data;
    }

    public static ObjectData HandlePolyline(Polyline polyline, ObjectData data, Point3d origin)
    {
      var polylineData = new PolylineData
      {
        Layer = polyline.Layer,
        Vectors = new List<SimpleVector3d>(),
        LineType = polyline.Linetype,
        Closed = polyline.Closed,
        StartWidth = polyline.GetStartWidthAt(0),
        EndWidth = polyline.GetEndWidthAt(polyline.NumberOfVertices - 1),
        GlobalWidth = polyline.ConstantWidth,
      };

      for (int i = 0; i < polyline.NumberOfVertices; i++)
      {
        Point3d point = polyline.GetPoint3dAt(i);
        Vector3d vector = point - origin;
        polylineData.Vectors.Add(
          new SimpleVector3d
          {
            X = vector.X,
            Y = vector.Y,
            Z = vector.Z,
          }
        );
        polylineData.Bulges.Add(polyline.GetBulgeAt(i));
      }

      data.Polylines.Add(polylineData);

      return data;
    }

    public static ObjectData HandlePolyline2d(Polyline2d polyline, ObjectData data, Point3d origin)
    {
      var polylineData = new Polyline2dData
      {
        Vertices = new List<SimpleVector3d>(),
        Bulges = new List<double>(),
        StartWidths = new List<double>(),
        EndWidths = new List<double>(),
        Closed = polyline.Closed,
        Layer = polyline.Layer,
      };

      foreach (ObjectId vertexId in polyline)
      {
        Vertex2d vertex = (Vertex2d)
          polyline.Database.TransactionManager.GetObject(vertexId, OpenMode.ForRead);
        polylineData.Vertices.Add(
          new SimpleVector3d
          {
            X = vertex.Position.X - origin.X,
            Y = vertex.Position.Y - origin.Y,
            Z = vertex.Position.Z - origin.Z,
          }
        );
        polylineData.Bulges.Add(vertex.Bulge);
        polylineData.StartWidths.Add(vertex.StartWidth);
        polylineData.EndWidths.Add(vertex.EndWidth);
      }

      data.Polylines2d.Add(polylineData);
      return data;
    }

    public static ObjectData HandleSolid(Solid solid, ObjectData data, Point3d origin)
    {
      var solidData = new SolidData { Layer = solid.Layer, Vertices = new List<SimpleVector3d>() };

      for (short i = 0; i < 4; i++)
      {
        Point3d point = solid.GetPointAt(i);
        Vector3d vector = point - origin;
        solidData.Vertices.Add(
          new SimpleVector3d
          {
            X = vector.X,
            Y = vector.Y,
            Z = vector.Z,
          }
        );
      }

      data.Solids.Add(solidData);

      return data;
    }

    public static ObjectData HandleText(DBText text, ObjectData data, Point3d origin)
    {
      if (text is AttributeDefinition attDef)
      {
        var attDefData = new AttributeDefinitionData
        {
          Layer = attDef.Layer,
          Style = attDef.TextStyleName,
          Justification = attDef.Justify,
          Contents = attDef.TextString,
          Tag = attDef.Tag,
          Prompt = attDef.Prompt,
          Location = new SimpleVector3d
          {
            X = attDef.Position.X - origin.X,
            Y = attDef.Position.Y - origin.Y,
            Z = attDef.Position.Z - origin.Z,
          },
          LineSpaceDistance = attDef.WidthFactor,
          Height = attDef.Height,
          Rotation = attDef.Rotation,
          AlignmentPoint = new SimpleVector3d
          {
            X = attDef.AlignmentPoint.X - origin.X,
            Y = attDef.AlignmentPoint.Y - origin.Y,
            Z = attDef.AlignmentPoint.Z - origin.Z,
          },
          HorizontalMode = attDef.HorizontalMode,
          IsMirroredInX = attDef.IsMirroredInX,
          IsMirroredInY = attDef.IsMirroredInY,
          Invisible = attDef.Invisible,
          StyleAttributes = new StyleAttributes
          {
            Height = attDef.Height,
            WidthFactor = attDef.WidthFactor,
            FontName = attDef.TextStyleName,
          },
        };

        data.AttributeDefinitions.Add(attDefData);
      }
      else
      {
        var textData = new TextData
        {
          Layer = text.Layer,
          Style = text.TextStyleName,
          Contents = text.TextString,
          Location = new SimpleVector3d
          {
            X = text.Position.X - origin.X,
            Y = text.Position.Y - origin.Y,
            Z = text.Position.Z - origin.Z,
          },
          LineSpaceDistance = text.WidthFactor,
          Height = text.Height,
          Rotation = text.Rotation,
          AlignmentPoint = new SimpleVector3d
          {
            X = text.AlignmentPoint.X - origin.X,
            Y = text.AlignmentPoint.Y - origin.Y,
            Z = text.AlignmentPoint.Z - origin.Z,
          },
          Justification = text.Justify,
          HorizontalMode = text.HorizontalMode,
          IsMirroredInX = text.IsMirroredInX,
          IsMirroredInY = text.IsMirroredInY,
          StyleAttributes = new StyleAttributes
          {
            Height = text.Height,
            WidthFactor = text.WidthFactor,
            FontName = text.TextStyleName,
          },
        };

        data.Texts.Add(textData);
      }

      return data;
    }

    public static ObjectData HandleHatch(Hatch hatch, ObjectData data, Point3d origin)
    {
      HatchData hatchData = HatchBoundary(hatch);

      Console.WriteLine("Polylines:");
      foreach (var polyline in hatchData.Polylines)
      {
        for (int i = 0; i < polyline.Vectors.Count; i++)
        {
          polyline.Vectors[i] = new SimpleVector3d
          {
            X = polyline.Vectors[i].X - origin.X,
            Y = polyline.Vectors[i].Y - origin.Y,
            Z = polyline.Vectors[i].Z - origin.Z,
          };
          Console.WriteLine($"\nPolyline");
          Console.WriteLine($"\nX:{polyline.Vectors[i].X}");
          Console.WriteLine($"\nY: {polyline.Vectors[i].Y}");
          Console.WriteLine($"\nZ: {polyline.Vectors[i].Z}");
        }
      }

      foreach (var line in hatchData.Lines)
      {
        line.StartPoint = new SimpleVector3d
        {
          X = line.StartPoint.X - origin.X,
          Y = line.StartPoint.Y - origin.Y,
          Z = line.StartPoint.Z - origin.Z,
        };

        line.EndPoint = new SimpleVector3d
        {
          X = line.EndPoint.X - origin.X,
          Y = line.EndPoint.Y - origin.Y,
          Z = line.EndPoint.Z - origin.Z,
        };
      }

      foreach (var arc in hatchData.Arcs)
      {
        arc.Center = new SimpleVector3d
        {
          X = arc.Center.X - origin.X,
          Y = arc.Center.Y - origin.Y,
          Z = arc.Center.Z - origin.Z,
        };
      }

      foreach (var circle in hatchData.Circles)
      {
        circle.Center = new SimpleVector3d
        {
          X = circle.Center.X - origin.X,
          Y = circle.Center.Y - origin.Y,
          Z = circle.Center.Z - origin.Z,
        };
      }

      foreach (var ellipse in hatchData.Ellipses)
      {
        ellipse.Center = new SimpleVector3d
        {
          X = ellipse.Center.X - origin.X,
          Y = ellipse.Center.Y - origin.Y,
          Z = ellipse.Center.Z - origin.Z,
        };
      }

      foreach (var spline in hatchData.Splines)
      {
        for (int i = 0; i < spline.ControlPoints.Count; i++)
        {
          spline.ControlPoints[i] = new SimpleVector3d
          {
            X = spline.ControlPoints[i].X - origin.X,
            Y = spline.ControlPoints[i].Y - origin.Y,
            Z = spline.ControlPoints[i].Z - origin.Z,
          };
        }

        for (int i = 0; i < spline.FitPoints.Count; i++)
        {
          spline.FitPoints[i] = new SimpleVector3d
          {
            X = spline.FitPoints[i].X - origin.X,
            Y = spline.FitPoints[i].Y - origin.Y,
            Z = spline.FitPoints[i].Z - origin.Z,
          };
        }
      }

      data.Hatches.Add(hatchData);

      return data;
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

    public static void SaveDataToJsonFile(object data, string filePath)
    {
      var objectData = (ObjectData)data;
      objectData.Scale = Scale;

      string jsonData = JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented);
      File.WriteAllText(filePath, jsonData);
    }
  }

  public class TextStyle
  {
    private StyleAttributes styleAttributes;

    public TextStyle(StyleAttributes styleAttributes)
    {
      this.styleAttributes = styleAttributes;
    }

    public void CreateStyleIfNotExisting(string name)
    {
      Autodesk.AutoCAD.ApplicationServices.Document acDoc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database acCurDb = acDoc.Database;

      using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
      {
        TextStyleTable acTextStyleTable;
        acTextStyleTable =
          acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;

        if (acTextStyleTable.Has(name) == false)
        {
          acTextStyleTable.UpgradeOpen();

          TextStyleTableRecord acTextStyleTableRec;
          acTextStyleTableRec = new TextStyleTableRecord();

          acTextStyleTableRec.Name = name;
          acTextStyleTableRec.FileName = styleAttributes.FontName;
          acTextStyleTableRec.TextSize = styleAttributes.Height;
          acTextStyleTableRec.XScale = styleAttributes.WidthFactor;

          acTextStyleTable.Add(acTextStyleTableRec);
          acTrans.AddNewlyCreatedDBObject(acTextStyleTableRec, true);
        }

        acTrans.Commit();
      }
    }

    public List<string> GetStyles()
    {
      List<string> styles = new List<string>();

      Autodesk.AutoCAD.ApplicationServices.Document acDoc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database acCurDb = acDoc.Database;

      using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
      {
        TextStyleTable acTextStyleTable;
        acTextStyleTable =
          acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;

        foreach (var styleId in acTextStyleTable)
        {
          TextStyleTableRecord acTextStyleTableRec;
          acTextStyleTableRec =
            acTrans.GetObject(styleId, OpenMode.ForRead) as TextStyleTableRecord;

          styles.Add(acTextStyleTableRec.Name);
        }

        acTrans.Commit();
      }

      return styles;
    }
  }

  public class HatchDataConverter : JsonConverter<HatchData>
  {
    public override HatchData ReadJson(
      JsonReader reader,
      Type objectType,
      HatchData existingValue,
      bool hasExistingValue,
      JsonSerializer serializer
    )
    {
      JObject jsonObject = JObject.Load(reader);

      HatchData hatchData = new HatchData
      {
        Layer = jsonObject["Layer"].ToString(),
        PatternName = jsonObject["PatternName"]?.ToString(),
        PatternScale = jsonObject["PatternScale"].ToObject<double>(),
        PatternType = jsonObject["PatternType"].ToObject<HatchPatternType>(),
        Angle = jsonObject["Angle"].ToObject<double>(),
        Polylines = jsonObject["Polylines"].ToObject<List<PolylineData>>(),
        Lines = jsonObject["Lines"].ToObject<List<LineData>>(),
        Arcs = jsonObject["Arcs"].ToObject<List<ArcData>>(),
        Circles = jsonObject["Circles"].ToObject<List<CircleData>>(),
        Ellipses = jsonObject["Ellipses"].ToObject<List<EllipseData>>(),
        Splines = jsonObject["Splines"].ToObject<List<SplineData>>(),
      };

      return hatchData;
    }

    public override void WriteJson(JsonWriter writer, HatchData value, JsonSerializer serializer)
    {
      writer.WriteStartObject();
      writer.WritePropertyName("Layer");
      writer.WriteValue(value.Layer);
      writer.WritePropertyName("PatternName");
      writer.WriteValue(value.PatternName);
      writer.WritePropertyName("PatternScale");
      writer.WriteValue(value.PatternScale);
      writer.WritePropertyName("PatternType");
      writer.WriteValue(value.PatternType);
      writer.WritePropertyName("Angle");
      writer.WriteValue(value.Angle);
      writer.WritePropertyName("Polylines");
      serializer.Serialize(writer, value.Polylines);
      writer.WritePropertyName("Lines");
      serializer.Serialize(writer, value.Lines);
      writer.WritePropertyName("Arcs");
      serializer.Serialize(writer, value.Arcs);
      writer.WritePropertyName("Circles");
      serializer.Serialize(writer, value.Circles);
      writer.WritePropertyName("Ellipses");
      serializer.Serialize(writer, value.Ellipses);
      writer.WritePropertyName("Splines");
      serializer.Serialize(writer, value.Splines);
      writer.WriteEndObject();
    }
  }

  public class ObjectData
  {
    public List<PolylineData> Polylines { get; set; }
    public List<Polyline2dData> Polylines2d { get; set; } // Add this line
    public List<LineData> Lines { get; set; }
    public List<ArcData> Arcs { get; set; }
    public List<CircleData> Circles { get; set; }
    public List<EllipseData> Ellipses { get; set; }
    public List<MTextData> MTexts { get; set; }
    public List<TextData> Texts { get; set; }
    public List<SolidData> Solids { get; set; }
    public List<HatchData> Hatches { get; set; }
    public List<SplineData> Splines { get; set; }
    public List<AttributeDefinitionData> AttributeDefinitions { get; set; }
    public int NumberOfRows { get; set; }
    public double Scale { get; set; }

    public ObjectData()
    {
      Polylines = new List<PolylineData>();
      Polylines2d = new List<Polyline2dData>();
      Lines = new List<LineData>();
      Arcs = new List<ArcData>();
      Circles = new List<CircleData>();
      Ellipses = new List<EllipseData>();
      MTexts = new List<MTextData>();
      Texts = new List<TextData>();
      Solids = new List<SolidData>();
      Hatches = new List<HatchData>();
      Splines = new List<SplineData>();

      AttributeDefinitions = new List<AttributeDefinitionData>();
    }
  }

  public class AttributeDefinitionData : TextData
  {
    public string Prompt { get; set; }
    public bool Invisible { get; set; }
  }

  public class TextData : BaseData
  {
    public string Style { get; set; }
    public AttachmentPoint Justification { get; set; }
    public string Contents { get; set; }
    public string Tag { get; set; }
    public SimpleVector3d Location { get; set; }
    public double LineSpaceDistance { get; set; }
    public double Height { get; set; }
    public double Rotation { get; set; }
    public SimpleVector3d AlignmentPoint { get; set; }
    public TextHorizontalMode HorizontalMode { get; set; }
    public bool IsMirroredInX { get; set; }
    public bool IsMirroredInY { get; set; }
    public StyleAttributes StyleAttributes { get; set; }
  }

  public class PolylineData : BaseData
  {
    public List<SimpleVector3d> Vectors { get; set; }
    public List<double> Bulges { get; set; }
    public string LineType { get; set; }
    public bool Closed { get; set; }
    public double StartWidth { get; set; }
    public double EndWidth { get; set; }
    public double GlobalWidth { get; set; }

    public PolylineData()
    {
      Vectors = new List<SimpleVector3d>();
      Bulges = new List<double>();
    }
  }

  public class SplineData : BaseData
  {
    public List<SimpleVector3d> ControlPoints { get; set; }
    public List<SimpleVector3d> FitPoints { get; set; }
    public List<double> Knots { get; set; }
    public List<double> Weights { get; set; }
    public int Degree { get; set; }
    public bool Rational { get; set; }
    public bool Closed { get; set; }
    public bool Periodic { get; set; }
    public SimpleVector3d StartTangent { get; set; }
    public SimpleVector3d EndTangent { get; set; }
    public double FitTolerance { get; set; }
    public double KnotTolerance { get; set; }

    public SplineData()
    {
      ControlPoints = new List<SimpleVector3d>();
      FitPoints = new List<SimpleVector3d>();
      Knots = new List<double>();
      Weights = new List<double>();
    }
  }

  public class LineData : BaseData
  {
    public SimpleVector3d StartPoint { get; set; }
    public SimpleVector3d EndPoint { get; set; }
  }

  public class ArcData : BaseData
  {
    public SimpleVector3d Center { get; set; }
    public double Radius { get; set; }
    public double StartAngle { get; set; }
    public double EndAngle { get; set; }
  }

  public class CircleData : BaseData
  {
    public SimpleVector3d Center { get; set; }
    public double Radius { get; set; }
  }

  public class EllipseData : BaseData
  {
    public SimpleVector3d UnitNormal { get; set; }
    public SimpleVector3d Center { get; set; }
    public SimpleVector3d MajorAxis { get; set; }
    public double MajorRadius { get; set; }
    public double MinorRadius { get; set; }
    public double StartAngle { get; set; }
    public double EndAngle { get; set; }

    public double RadiusRatio()
    {
      if (MinorRadius != 0 && MajorRadius != 0)
      {
        return MinorRadius / MajorRadius;
      }
      else
      {
        return 0;
      }
    }
  }

  [JsonConverter(typeof(HatchDataConverter))]
  public class HatchData : BaseData
  {
    public double PatternScale { get; set; }
    public double Angle { get; set; }
    public string PatternName { get; internal set; }
    public HatchPatternType PatternType { get; internal set; }
    public List<PolylineData> Polylines { get; set; }
    public List<LineData> Lines { get; set; }
    public List<ArcData> Arcs { get; set; }
    public List<CircleData> Circles { get; set; }
    public List<EllipseData> Ellipses { get; set; }
    public List<SplineData> Splines { get; set; }

    public HatchData()
    {
      Polylines = new List<PolylineData>();
      Lines = new List<LineData>();
      Arcs = new List<ArcData>();
      Circles = new List<CircleData>();
      Ellipses = new List<EllipseData>();
      Splines = new List<SplineData>();
    }
  }

  public class MTextData : BaseData
  {
    public string Style { get; set; }
    public string Justification { get; set; }
    public string Contents { get; set; }
    public Vector3d Direction { get; set; }
    public SimpleVector3d Location { get; set; }
    public double LineSpaceDistance { get; set; }
    public double LineSpaceFactor { get; set; }
    public LineSpacingStyle LineSpacingStyle { get; set; }
    public double TextHeight { get; set; }
    public double Width { get; set; }
    public double Rotation { get; set; }
    public StyleAttributes StyleAttributes { get; set; }
  }

  public class SolidData : BaseData
  {
    public List<SimpleVector3d> Vertices { get; set; }
  }

  public class BaseData
  {
    public string Layer { get; set; }
  }

  public class SimpleVector3d
  {
    public SimpleVector3d(double X = 0, double Y = 0, double Z = 0)
    {
      this.X = X;
      this.Y = Y;
      this.Z = Z;
    }

    public double X { get; set; }
    public double Y { get; set; }
    public double Z { get; set; }
  }

  public class Polyline2dData : BaseData
  {
    public List<SimpleVector3d> Vertices { get; set; }
    public List<double> Bulges { get; set; } // Add bulges
    public List<double> StartWidths { get; set; } // Add start widths
    public List<double> EndWidths { get; set; } // Add end widths
    public bool Closed { get; set; }

    public Polyline2dData()
    {
      Vertices = new List<SimpleVector3d>();
      Bulges = new List<double>();
      StartWidths = new List<double>();
      EndWidths = new List<double>();
    }
  }

  public class StyleAttributes
  {
    public double Height { get; set; }
    public double WidthFactor { get; set; }
    public string FontName { get; set; }
  }

  public class GeneralTextJig : EntityJig
  {
    private Point3d insertionPoint;
    private string textString;
    private double textHeight;
    private TextHorizontalMode horizontalMode;

    public Point3d InsertionPoint => insertionPoint;

    public GeneralTextJig(string textString, double textHeight, TextHorizontalMode horizontalMode)
      : base(new DBText())
    {
      this.textString = textString;
      this.textHeight = textHeight;
      this.horizontalMode = horizontalMode;

      DBText dbText = (DBText)Entity;
      dbText.TextString = textString;
      dbText.Height = textHeight;
      dbText.HorizontalMode = horizontalMode;
    }

    protected override bool Update()
    {
      DBText dbText = (DBText)Entity;
      dbText.Position = insertionPoint;

      if (horizontalMode != TextHorizontalMode.TextLeft)
      {
        dbText.AlignmentPoint = insertionPoint;
      }

      return true;
    }

    protected override SamplerStatus Sampler(JigPrompts prompts)
    {
      JigPromptPointOptions jigOpts = new JigPromptPointOptions("\nSpecify insertion point: ");
      PromptPointResult ppr = prompts.AcquirePoint(jigOpts);

      if (ppr.Status == PromptStatus.OK)
      {
        if (ppr.Value.IsEqualTo(insertionPoint))
        {
          return SamplerStatus.NoChange;
        }
        else
        {
          insertionPoint = ppr.Value;
          return SamplerStatus.OK;
        }
      }

      return SamplerStatus.Cancel;
    }
  }
}
