using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ElectricalCommands
{
  internal class CADtoFormImport
  {
    [CommandMethod("CADPANELIMPORT")]
    public void CADPANELIMPORT()
    {
      Document doc = Autodesk
          .AutoCAD
          .ApplicationServices
          .Application
          .DocumentManager
          .MdiActiveDocument;
      Editor ed = doc.Editor;

      // Prompt the user to select objects
      PromptSelectionResult selection = ed.GetSelection();

      if (selection.Status != PromptStatus.OK)
      {
        return;
      }

      using (Transaction tr = doc.TransactionManager.StartTransaction())
      {
        Dictionary<string, Point3d> selectionPositions = new Dictionary<string, Point3d>();
        List<Dictionary<string, Point3d>> panelSelectionAreas =
            new List<Dictionary<string, Point3d>>();

        foreach (SelectedObject selectedObject in selection.Value)
        {
          Entity entity = (Entity)tr.GetObject(selectedObject.ObjectId, OpenMode.ForRead);

          if (entity is DBText dbText)
          {
            if (
                dbText.TextString.Contains("PANEL")
                && !dbText.TextString.Contains("LOAD")
                && dbText.Height >= 0.12
            )
            {
              Point3d pt_1 = new Point3d(
                  dbText.Position.X,
                  dbText.Position.Y + 0.3744,
                  0
              );
              Point3d pt_2 = new Point3d(dbText.Position.X + 1, dbText.Position.Y, 0);

              PromptSelectionResult selection_2 = ed.SelectCrossingWindow(pt_1, pt_2);

              if (selection_2.Status == PromptStatus.OK)
              {
                foreach (SelectedObject thing in selection_2.Value)
                {
                  using (
                      Transaction tr2 = doc.TransactionManager.StartTransaction()
                  )
                  {
                    Entity sub_entity = (Entity)
                        tr2.GetObject(thing.ObjectId, OpenMode.ForRead);
                    if (sub_entity is Polyline polyline)
                    {
                      Point3d topLeft = new Point3d(
                          polyline.Bounds.Value.MinPoint.X,
                          polyline.Bounds.Value.MaxPoint.Y,
                          0
                      );
                      Point3d bottomRight = new Point3d(
                          polyline.Bounds.Value.MaxPoint.X,
                          polyline.Bounds.Value.MinPoint.Y,
                          0
                      );

                      selectionPositions.Add("top_left", topLeft);
                      selectionPositions.Add("bottom_right", bottomRight);
                    }
                  }
                }
              }
            }
          }
          else if (entity is MText mText)
          {
            if (
                mText.Contents.Contains("PANEL")
                && !mText.Contents.Contains("LOAD")
                && mText.TextHeight >= 0.12
            )
            {
              Point3d pt_1 = new Point3d(
                  mText.Location.X,
                  mText.Location.Y + 0.3744,
                  0
              );
              Point3d pt_2 = new Point3d(mText.Location.X + 1, mText.Location.Y, 0);

              PromptSelectionResult selection_2 = ed.SelectCrossingWindow(pt_1, pt_2);

              if (selection_2.Status == PromptStatus.OK)
              {
                foreach (SelectedObject thing in selection_2.Value)
                {
                  using (
                      Transaction tr2 = doc.TransactionManager.StartTransaction()
                  )
                  {
                    Entity sub_entity = (Entity)
                        tr2.GetObject(thing.ObjectId, OpenMode.ForRead);
                    if (sub_entity is Line line)
                    {
                      Point3d topLeft = new Point3d(
                          line.Bounds.Value.MinPoint.X,
                          line.Bounds.Value.MaxPoint.Y,
                          0
                      );

                      if (!selectionPositions.ContainsKey("top_left"))
                      {
                        selectionPositions.Add("top_left", topLeft);

                        MText subtotal = LocateSubTotal(line);

                        selectionPositions = GetBottomRightCoordinate(
                            doc,
                            ed,
                            selectionPositions,
                            subtotal
                        );
                      }
                    }
                  }
                }
              }
            }
          }
          if (
              selectionPositions.ContainsKey("top_left")
              && selectionPositions.ContainsKey("bottom_right")
          )
          {
            panelSelectionAreas.Add(selectionPositions);
            selectionPositions = new Dictionary<string, Point3d>();
          }
        }

        PanelBreakdown(panelSelectionAreas);
      }
    }

    [CommandMethod("SETPANELREGION")]
    public void SetPanelRegion()
    {
      var doc = Autodesk
          .AutoCAD
          .ApplicationServices
          .Application
          .DocumentManager
          .MdiActiveDocument;
      var ed = doc.Editor;

      // Prompt for the origin point
      PromptPointResult ppr = ed.GetPoint("\nEnter the origin point: ");
      if (ppr.Status != PromptStatus.OK)
        return;
      Point3d origin = ppr.Value;

      // Prompt for the first corner point
      ppr = ed.GetPoint("\nEnter the first corner point: ");
      if (ppr.Status != PromptStatus.OK)
        return;
      Point3d pt1 = ppr.Value;

      // Prompt for the second corner point
      ppr = ed.GetPoint("\nEnter the second corner point: ");
      if (ppr.Status != PromptStatus.OK)
        return;
      Point3d pt2 = ppr.Value;

      // Define the selection filter
      TypedValue[] filter = new TypedValue[]
      {
                new TypedValue((int)DxfCode.Start, "TEXT,MTEXT"),
      };

      // Select the text and MText entities within the rectangular region
      PromptSelectionResult psr = ed.SelectCrossingWindow(
          pt1,
          pt2,
          new SelectionFilter(filter)
      );
      if (psr.Status != PromptStatus.OK)
        return;

      // Prompt for the selection name
      PromptResult pr = ed.GetString("\nEnter a name for the selection: ");
      if (pr.Status != PromptStatus.OK)
        return;
      string name = pr.StringResult;

      // Get the text and MText entities
      List<string> text = new List<string>();
      using (Transaction tr = doc.TransactionManager.StartTransaction())
      {
        foreach (SelectedObject so in psr.Value)
        {
          var entity = tr.GetObject(so.ObjectId, OpenMode.ForRead);

          if (entity is DBText dbText)
          {
            text.Add(dbText.TextString);
          }
          else if (entity is MText mText)
          {
            text.Add(mText.Contents);
          }
        }
      }

      // Create the strings
      string point1 =
          $"var pt1 = new Dictionary<string, object> {{ {{ \"x\", {pt1.X - origin.X} }}, {{ \"y\", {pt1.Y - origin.Y} }} }};";
      string point2 =
          $"var pt2 = new Dictionary<string, object> {{ {{ \"x\", {pt2.X - origin.X} }}, {{ \"y\", {pt2.Y - origin.Y} }} }};";

      // Combine the strings
      string combined = $"{point1}\n{point2}\n";
      foreach (var item in text)
      {
        combined += $"{item}\n";
      }

      // Save the string to a file
      string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
      string path = Path.Combine(desktopPath, name + ".txt");
      File.WriteAllText(path, combined);
    }

    private static Dictionary<string, Point3d> GetBottomRightCoordinate(Document doc, Editor ed, Dictionary<string, Point3d> selectionPositions, MText mText)
    {
      Point3d pt_1 = new Point3d(mText.Location.X, mText.Location.Y - 0.3744, 0);
      Point3d pt_2 = new Point3d(mText.Location.X + 1, mText.Location.Y, 0);

      PromptSelectionResult selection_2 = ed.SelectCrossingWindow(pt_1, pt_2);

      if (selection_2.Status == PromptStatus.OK)
      {
        foreach (SelectedObject thing in selection_2.Value)
        {
          using (Transaction tr2 = doc.TransactionManager.StartTransaction())
          {
            Entity sub_entity = (Entity)tr2.GetObject(thing.ObjectId, OpenMode.ForRead);
            if (sub_entity is Line line)
            {
              if (line.Length > 8)
              {
                Point3d bottomRight = new Point3d(
                    line.Bounds.Value.MaxPoint.X,
                    line.Bounds.Value.MinPoint.Y,
                    0
                );

                if (!selectionPositions.ContainsKey("bottom_right"))
                {
                  selectionPositions.Add("bottom_right", bottomRight);
                  return selectionPositions;
                }
              }
            }
          }
        }
      }
      return selectionPositions;
    }

    private MText LocateSubTotal(Line line)
    {
      var (doc, db, ed) = GetGlobals();

      using (var tr = db.TransactionManager.StartTransaction())
      {
        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        var btr = (BlockTableRecord)
            tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForRead);

        double minDistance = double.MaxValue;
        MText closestText = null;

        foreach (var id in btr)
        {
          var entity = (Entity)tr.GetObject(id, OpenMode.ForRead);
          if (entity is MText text)
          {
            if (
                text.Location.Y < line.StartPoint.Y
                && text.Location.X >= line.StartPoint.X
                && text.Location.X <= line.EndPoint.X
                && text.Contents == "\\FArial; SUB-TOTAL"
            )
            {
              var distance = line.StartPoint.Y - text.Location.Y;
              if (distance < minDistance)
              {
                minDistance = distance;
                closestText = text;
              }
            }
          }
        }
        tr.Commit();
        return closestText;
      }
    }

    public void PanelBreakdown(List<Dictionary<string, Point3d>> panelSelectionAreas)
    {
      var resultList = new List<Dictionary<string, object>>();
      var doc = Autodesk
          .AutoCAD
          .ApplicationServices
          .Application
          .DocumentManager
          .MdiActiveDocument;
      var ed = doc.Editor;

      using (Transaction tr = doc.TransactionManager.StartTransaction())
      {
        ZoomCamera(panelSelectionAreas);

        foreach (var area in panelSelectionAreas)
        {
          if (area.ContainsKey("top_left") && area.ContainsKey("bottom_right"))
          {
            Point3d pt1 = area["top_left"];
            Point3d pt2 = area["bottom_right"];

            TypedValue[] filter = [new TypedValue((int)DxfCode.Start, "TEXT,MTEXT"),];

            PromptSelectionResult psr = ed.SelectCrossingWindow(
                pt1,
                pt2,
                new SelectionFilter(filter)
            );

            var textList = new List<Dictionary<string, object>>();

            foreach (SelectedObject so in psr.Value)
            {
              var entity = tr.GetObject(so.ObjectId, OpenMode.ForRead);

              string textValue = null;
              Point3d position = new Point3d();

              if (entity is DBText dbText)
              {
                textValue = dbText.TextString;
                position = dbText.Position;
              }
              else if (entity is MText mText)
              {
                textValue = mText.Contents;
                position = mText.Location;
              }

              if (textValue != null)
              {
                textValue = textValue
                    .Replace("\\FArial;", "")
                    .Replace("\\Farial|c0;", "")
                    .Replace("{", "")
                    .Replace("}", "")
                    .Trim();

                double relativeX = position.X - pt1.X;
                double relativeY = position.Y - pt1.Y;

                textList.Add(
                    new Dictionary<string, object>
                    {
                                        { "value", textValue },
                                        { "x", relativeX },
                                        { "y", relativeY }
                    }
                );
              }
            }

            var result = new Dictionary<string, object>
                        {
                            {
                                "polyline",
                                new Dictionary<string, Dictionary<string, object>>
                                {
                                    {
                                        "top_left",
                                        new Dictionary<string, object> { { "x", 0 }, { "y", 0 } }
                                    },
                                    {
                                        "top_right",
                                        new Dictionary<string, object>
                                        {
                                            { "x", pt2.X - pt1.X },
                                            { "y", 0 }
                                        }
                                    },
                                    {
                                        "bottom_right",
                                        new Dictionary<string, object>
                                        {
                                            { "x", pt2.X - pt1.X },
                                            { "y", pt2.Y - pt1.Y }
                                        }
                                    },
                                    {
                                        "bottom_left",
                                        new Dictionary<string, object>
                                        {
                                            { "x", 0 },
                                            { "y", pt2.Y - pt1.Y }
                                        }
                                    }
                                }
                            },
                            { "text", textList }
                        };

            resultList.Add(result);
          }
        }

        tr.Commit();
      }

      ParseCADPanelObjects(resultList);
    }

    private void ZoomCamera(List<Dictionary<string, Point3d>> panelSelectionAreas)
    {
      var doc = Autodesk
          .AutoCAD
          .ApplicationServices
          .Application
          .DocumentManager
          .MdiActiveDocument;
      var ed = doc.Editor;

      double minX = double.MaxValue;
      double minY = double.MaxValue;
      double maxX = double.MinValue;
      double maxY = double.MinValue;

      foreach (var area in panelSelectionAreas)
      {
        if (area.ContainsKey("top_left"))
        {
          Point3d topLeft = area["top_left"];
          minX = Math.Min(minX, topLeft.X);
          maxY = Math.Max(maxY, topLeft.Y);
        }

        if (area.ContainsKey("bottom_right"))
        {
          Point3d bottomRight = area["bottom_right"];
          maxX = Math.Max(maxX, bottomRight.X);
          minY = Math.Min(minY, bottomRight.Y);
        }
      }

      Point3d pt1 = new Point3d(minX, maxY, 0);
      Point3d pt2 = new Point3d(maxX, minY, 0);

      string cmd = string.Format(
          "._zoom _window {0},{1},{2} {3},{4},{5} ",
          pt1.X,
          pt1.Y,
          pt1.Z,
          pt2.X,
          pt2.Y,
          pt2.Z
      );

      doc.SendStringToExecute(cmd, true, false, true);
    }

    private void ParseCADPanelObjects(List<Dictionary<string, object>> resultList)
    {
      var parsedDataList = new List<Dictionary<string, object>>();

      foreach (var result in resultList)
      {
        var panelName = ParsePanelName(result);

        if (panelName == "")
        {
          continue;
        }

        var location = ParseLocation(result);
        var main = ParseMain(result);
        var bus_rating = ParseBusRating(result);
        var voltage_low = ParseVoltageLow(result);
        var voltage_high = ParseVoltageHigh(result);
        var phase = ParsePhase(result);
        var wire = ParseWire(result);
        var mounting = ParseMounting(result);

        Dictionary<string, object> parsedData = new Dictionary<string, object>
                {
                    { "panelName", panelName },
                    { "location", location },
                    { "main", main },
                    { "bus_rating", bus_rating },
                    { "voltage_low", voltage_low },
                    { "voltage_high", voltage_high },
                    { "phase", phase },
                    { "wire", wire },
                    { "mounting", mounting }
                };

        parsedDataList.Add(parsedData);
      }
    }

    private object ParseMounting(Dictionary<string, object> result)
    {
      var pt1 = new Dictionary<string, object>
            {
                { "x", 6.67500000000001 },
                { "y", -0.200000000000003 }
            };
      var pt2 = new Dictionary<string, object>
            {
                { "x", 8.98559999999998 },
                { "y", -0.400000000000006 }
            };

      var textValues = GetTextValuesInRegion(result, pt1, pt2);
      var mounting = string.Join("", textValues).Replace(" ", "");

      mounting = mounting.Replace("MOUNTING", "").Replace(":", "").Replace(" ", "");

      if (mounting == "REC" || mounting == "RECESS")
      {
        mounting = "RECESSED";
      }

      return mounting;
    }

    private object ParseWire(Dictionary<string, object> result)
    {
      var pt1 = new Dictionary<string, object> { { "x", 6.67500000000001 }, { "y", 0 } };
      var pt2 = new Dictionary<string, object>
            {
                { "x", 8.98559999999998 },
                { "y", -0.187199999999848 }
            };

      var textValues = GetTextValuesInRegion(result, pt1, pt2);
      var wire = string.Join("", textValues).Replace(" ", "");
      var index = wire.IndexOf('W');

      if (index > 0)
      {
        return wire[index - 1];
      }
      else
      {
        return "3";
      }
    }

    private object ParsePhase(Dictionary<string, object> result)
    {
      var pt1 = new Dictionary<string, object> { { "x", 6.67500000000001 }, { "y", 0 } };
      var pt2 = new Dictionary<string, object>
            {
                { "x", 8.98559999999998 },
                { "y", -0.187199999999848 }
            };

      var textValues = GetTextValuesInRegion(result, pt1, pt2);
      var phase = string.Join("", textValues).Replace(" ", "");
      var index = phase.IndexOf('V');

      if (index < phase.Length - 1)
      {
        return phase[index + 1];
      }
      else
      {
        return "1";
      }
    }

    private object ParseVoltageHigh(Dictionary<string, object> result)
    {
      var pt1 = new Dictionary<string, object> { { "x", 6.67500000000001 }, { "y", 0 } };
      var pt2 = new Dictionary<string, object>
            {
                { "x", 8.98559999999998 },
                { "y", -0.187199999999848 }
            };

      var textValues = GetTextValuesInRegion(result, pt1, pt2);

      foreach (var textValue in textValues)
      {
        if (textValue.Contains("208"))
        {
          return "208";
        }
        else if (textValue.Contains("240"))
        {
          return "240";
        }
        else if (textValue.Contains("480"))
        {
          return "480";
        }
      }
      return "208";
    }

    private object ParseVoltageLow(Dictionary<string, object> result)
    {
      var pt1 = new Dictionary<string, object> { { "x", 6.67500000000001 }, { "y", 0 } };
      var pt2 = new Dictionary<string, object>
            {
                { "x", 8.98559999999998 },
                { "y", -0.187199999999848 }
            };

      var textValues = GetTextValuesInRegion(result, pt1, pt2);

      foreach (var textValue in textValues)
      {
        if (textValue.Contains("120"))
        {
          return "120";
        }
        else if (textValue.Contains("277"))
        {
          return "277";
        }
      }
      return "120";
    }

    private string ParseBusRating(Dictionary<string, object> result)
    {
      var pt1 = new Dictionary<string, object> { { "x", 5.06661018521663 }, { "y", 0 } };
      var pt2 = new Dictionary<string, object> { { "x", 6.67 }, { "y", -0.374399999999696 } };
      var textValues = GetTextValuesInRegion(result, pt1, pt2);
      var main = string.Join(" ", textValues).ToUpper();
      return new string(main.Where(char.IsDigit).ToArray());
    }

    private string ParseLocation(Dictionary<string, object> result)
    {
      var pt1 = new Dictionary<string, object> { { "x", 2.2222 }, { "y", 0.0 } };
      var pt2 = new Dictionary<string, object>
            {
                { "x", 5.0666 },
                { "y", -0.1872000000000007 }
            };
      var textValues = GetTextValuesInRegion(result, pt1, pt2);
      var location = string.Join(" ", textValues).ToUpper();

      return location.Replace("LOCATION", "").Trim();
    }

    private string ParsePanelName(Dictionary<string, object> result)
    {
      var pt1 = new Dictionary<string, object> { { "x", 0.0 }, { "y", 0.0 } };
      var pt2 = new Dictionary<string, object>
            {
                { "x", 2.2222 },
                { "y", -0.37439999999999962 }
            };

      if (!result.ContainsKey("text"))
      {
        return "";
      }

      var textList = (List<Dictionary<string, object>>)result["text"];
      var textValues = GetTextValuesInRegion(result, pt1, pt2);
      var panelName = string.Join(" ", textValues);

      if (!panelName.Contains("PANEL"))
      {
        return "";
      }

      if (!panelName.Contains('\'') && !panelName.Contains('`'))
      {
        return "";
      }

      int firstQuoteIndex = panelName.IndexOfAny(['\'', '`']);
      int lastQuoteIndex = panelName.LastIndexOfAny(['\'', '`']);

      if (firstQuoteIndex >= 0 && lastQuoteIndex > firstQuoteIndex)
      {
        return panelName.Substring(
            firstQuoteIndex + 1,
            lastQuoteIndex - firstQuoteIndex - 1
        );
      }

      return panelName;
    }

    private string ParseMain(Dictionary<string, object> result)
    {
      var pt1 = new Dictionary<string, object>
            {
                { "x", 2.2221755722507623 },
                { "y", -0.18719999999984793 }
            };
      var pt2 = new Dictionary<string, object>
            {
                { "x", 5.0666101852166321 },
                { "y", -0.37439999999969586 }
            };
      var textValues = GetTextValuesInRegion(result, pt1, pt2);
      var main = string.Join(" ", textValues).ToUpper();
      var digitsOnly = new string(main.Where(char.IsDigit).ToArray());

      if (main.Contains("M") && main.Contains("L") && main.Contains("O"))
      {
        return "M.L.O.";
      }
      else
      {
        return digitsOnly;
      }
    }

    private string[] GetTextValuesInRegion(Dictionary<string, object> result, Dictionary<string, object> pt1, Dictionary<string, object> pt2)
    {
      var textList = ((List<Dictionary<string, object>>)result["text"])
          .OrderBy(text => (double)text["x"]) // Then sort by X-coordinate in ascending order (left to right)
          .ToList();

      var textValues = new List<string>();

      foreach (var text in textList)
      {
        var x = (double)text["x"];
        var y = (double)text["y"];

        if (pt1["x"].GetType() == typeof(int))
        {
          pt1["x"] = Convert.ToDouble(pt1["x"]);
        }
        if (pt1["y"].GetType() == typeof(int))
        {
          pt1["y"] = Convert.ToDouble(pt1["y"]);
        }
        if (pt2["x"].GetType() == typeof(int))
        {
          pt2["x"] = Convert.ToDouble(pt2["x"]);
        }
        if (pt2["y"].GetType() == typeof(int))
        {
          pt2["y"] = Convert.ToDouble(pt2["y"]);
        }

        if (
            x >= (double)pt1["x"]
            && x <= (double)pt2["x"]
            && y >= (double)pt2["y"]
            && y <= (double)pt1["y"]
        )
        {
          textValues.Add(text["value"].ToString());
        }
      }

      return textValues.ToArray();
    }

    public static (Document doc, Database db, Editor ed) GetGlobals()
    {
      var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      var db = doc.Database;
      var ed = doc.Editor;

      return (doc, db, ed);
    }
  }
}