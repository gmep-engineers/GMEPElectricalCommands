using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using OfficeOpenXml;

namespace ElectricalCommands
{
  public class PanelNote
  {
    public int Number;
    public string PanelId;
    public int CircuitNo;
    public int Length;
    public string Note;

    public PanelNote(int Number, string PanelId, int CircuitNo, int Length, string Note)
    {
      this.Number = Number;
      this.PanelId = PanelId;
      this.CircuitNo = CircuitNo;
      this.Length = Length;
      this.Note = Note;
    }
  }

  public class JsonPanel3P
  {
    public string id;
    public bool read_only;
    public string main;
    public string panel;
    public string location;
    public string voltage1;
    public string voltage2;
    public string phase;
    public string wire;
    public string mounting;
    public string existing;
    public bool lcl_override;
    public bool lml_override;
    public string subtotal_a;
    public string subtotal_b;
    public string subtotal_c;
    public string total_va;
    public double lcl;
    public double lcl125;
    public double lml;
    public double lml125;
    public string kva;
    public string feeder_amps;
    public string custom_title;
    public string bus_rating;
    public bool[] description_left_highlights;
    public bool[] description_right_highlights;
    public bool[] breaker_left_highlights;
    public bool[] breaker_right_highlights;
    public string[] description_left;
    public string[] description_right;
    public string[] phase_a_left;
    public string[] phase_a_right;
    public string[] phase_b_left;
    public string[] phase_b_right;
    public string[] phase_c_left;
    public string[] phase_c_right;
    public string[] breaker_left;
    public string[] breaker_right;
    public string[] circuit_left;
    public string[] circuit_right;
    public string[] phase_a_left_tag;
    public string[] phase_a_right_tag;
    public string[] phase_b_left_tag;
    public string[] phase_b_right_tag;
    public string[] phase_c_left_tag;
    public string[] phase_c_right_tag;
    public string[] description_left_tags;
    public string[] description_right_tags;
    public string[] notes;
  }

  public class JsonPanel2P
  {
    public string id;
    public bool read_only;
    public string main;
    public string panel;
    public string location;
    public string voltage1;
    public string voltage2;
    public string phase;
    public string wire;
    public string mounting;
    public string existing;
    public bool lcl_override;
    public bool lml_override;
    public string subtotal_a;
    public string subtotal_b;
    public string subtotal_c;
    public string total_va;
    public double lcl;
    public double lcl125;
    public double lml;
    public double lml125;
    public string kva;
    public string feeder_amps;
    public string custom_title;
    public string bus_rating;
    public bool[] description_left_highlights;
    public bool[] description_right_highlights;
    public bool[] breaker_left_highlights;
    public bool[] breaker_right_highlights;
    public string[] description_left;
    public string[] description_right;
    public string[] phase_a_left;
    public string[] phase_a_right;
    public string[] phase_b_left;
    public string[] phase_b_right;
    public string[] breaker_left;
    public string[] breaker_right;
    public string[] circuit_left;
    public string[] circuit_right;
    public string[] phase_a_left_tag;
    public string[] phase_a_right_tag;
    public string[] phase_b_left_tag;
    public string[] phase_b_right_tag;
    public string[] description_left_tags;
    public string[] description_right_tags;
    public string[] notes;
  }

  public class PanelCommands
  {
    private MainForm myForm;
    private int circuitNumLeft = 1;
    private int circuitNumRight = 2;

    [CommandMethod("PANEL")]
    public void PANEL()
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Editor ed = doc.Editor;

      try
      {
        if (this.myForm != null && !this.myForm.IsDisposed)
        {
          // Bring myForm to the front
          this.myForm.BringToFront();
        }
        else
        {
          // Create a new MainForm if it's not already open
          this.myForm = new MainForm(this);
          this.myForm.InitializeModal();
          this.myForm.Show();
        }
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage("Error: " + ex.ToString());
      }
    }

    [CommandMethod("CREATEBLOCK")]
    public void CREATEBLOCK()
    {
      var (doc, db, _) = PanelCommands.GetGlobals();

      using (Transaction tr = db.TransactionManager.StartTransaction())
      {
        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

        BlockTableRecord existingBtr = null;
        ObjectId existingBtrId = ObjectId.Null;

        // Check if block already exists
        if (bt.Has("CIRCLEI"))
        {
          existingBtrId = bt["CIRCLEI"];

          if (existingBtrId != ObjectId.Null)
          {
            existingBtr = (BlockTableRecord)tr.GetObject(existingBtrId, OpenMode.ForWrite);

            if (existingBtr != null && existingBtr.Name == "CIRCLEI")
            {
              return; // Exit the function if existing block matches the new block
            }
          }
        }

        // Delete existing block and its contents
        if (existingBtr != null)
        {
          foreach (ObjectId id in existingBtr.GetBlockReferenceIds(true, true))
          {
            DBObject obj = tr.GetObject(id, OpenMode.ForWrite);
            obj.Erase(true);
          }

          existingBtr.Erase(true);
        }

        BlockTableRecord btr = new BlockTableRecord();
        btr.Name = "CIRCLEI";

        bt.UpgradeOpen();
        ObjectId btrId = bt.Add(btr);
        tr.AddNewlyCreatedDBObject(btr, true);

        // Create a circle centered at 0,0 with radius 2.0
        Circle circle = new Circle(new Point3d(0, 0, 0), new Vector3d(0, 0, 1), 0.09);
        circle.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 2); // Set circle color to yellow

        btr.AppendEntity(circle);
        tr.AddNewlyCreatedDBObject(circle, true);

        // Create a text entity
        DBText text = new DBText();
        text.Position = new Point3d(-0.042, -0.045, 0); // centered at the origin
        text.Height = 0.09; // Set the text height
        text.TextString = "1";
        text.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 2); // Set text color to yellow

        // Check if the text style "gmep" exists
        TextStyleTable textStyleTable = (TextStyleTable)
          tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
        if (textStyleTable.Has("gmep"))
        {
          text.TextStyleId = textStyleTable["gmep"]; // apply the "gmep" text style to the text entity
        }

        // Check if the layer "E-TEXT" exists
        LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
        if (lt.Has("E-TEXT"))
        {
          circle.Layer = "E-TEXT"; // Set the layer of the circle to "E-TEXT"
          text.Layer = "E-TEXT"; // Set the layer of the text to "E-TEXT"
        }

        btr.AppendEntity(text);
        tr.AddNewlyCreatedDBObject(text, true);

        tr.Commit();
      }
    }

    [CommandMethod("KEEPBREAKERS")]
    public void KEEPBREAKERS()
    {
      var (doc, db, ed) = PanelCommands.GetGlobals();

      using (var tr = db.TransactionManager.StartTransaction())
      {
        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

        if (!bt.Has("CIRCLEI"))
        {
          PromptKeywordOptions pko = new PromptKeywordOptions(
            "\nThe block 'CIRCLEI' does not exist. Do you want to create it? [Yes/No] ",
            "Yes No"
          );
          pko.AllowNone = true;
          PromptResult pr = ed.GetKeywords(pko);
          String prompt = pr.StringResult.ToLower();
          if (prompt == "no" || prompt == "n")
            return;
          else if (prompt == "yes" || prompt == "y")
            CREATEBLOCK();
        }

        var btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForWrite);

        var point1 = ed.GetPoint("\nSelect first point: ").Value;
        var point2 = ed.GetPoint("\nSelect second point: ").Value;

        // Swap the points if the y-coordinate of the first point is lower than that of the second point
        if (point1.Y < point2.Y)
        {
          Point3d tempPoint = point1;
          point1 = point2;
          point2 = tempPoint;
        }

        var point3 = ed.GetPoint("\nSelect third point: ").Value;

        var direction = point3.X > point1.X ? 1 : -1;
        var dist = (point1 - point2).Length;

        var line1Start = new Point3d(point1.X + direction * 0.05, point1.Y, 0);
        var line1End = new Point3d(line1Start.X + direction * 0.2, line1Start.Y, 0);
        var line2Start = new Point3d(line1Start.X, point2.Y, 0);
        var line2End = new Point3d(line1End.X, line2Start.Y, 0);

        string layerName = CreateOrGetLayer("E-TEXT", db, tr);

        var line1 = new Line(line1Start, line1End) { Layer = layerName };
        var line2 = new Line(line2Start, line2End) { Layer = layerName };

        var mid1 = new Point3d((line1Start.X + line1End.X) / 2, (line1Start.Y + line1End.Y) / 2, 0);
        var mid2 = new Point3d((line2Start.X + line2End.X) / 2, (line2Start.Y + line2End.Y) / 2, 0);
        var mid3 = new Point3d((mid1.X + mid2.X) / 2, (mid1.Y + mid2.Y) / 2, 0);

        var circleTop = new Point3d(mid3.X, mid3.Y + 0.09, 0);
        var circleBottom = new Point3d(mid3.X, mid3.Y - 0.09, 0);

        var line3 = new Line(mid1, circleTop) { Layer = layerName };
        var line4 = new Line(mid2, circleBottom) { Layer = layerName };

        if (dist > 0.3)
        {
          btr.AppendEntity(line1);
          btr.AppendEntity(line2);
          tr.AddNewlyCreatedDBObject(line1, true);
          tr.AddNewlyCreatedDBObject(line2, true);
        }

        var blkRef = new BlockReference(mid3, bt["CIRCLEI"]) { Layer = layerName };
        btr.AppendEntity(blkRef);
        tr.AddNewlyCreatedDBObject(blkRef, true);

        if (dist > 0.3)
        {
          btr.AppendEntity(line3);
          btr.AppendEntity(line4);
          tr.AddNewlyCreatedDBObject(line3, true);
          tr.AddNewlyCreatedDBObject(line4, true);
        }

        tr.Commit();
      }
    }

    [CommandMethod("IMPORTPANELS")]
    public void IMPORTPANELS()
    {
      CreatePanels(null);
    }

    private void CreateTextsWithoutPanelData(
      Transaction tr,
      string layerName,
      Point3d startPoint,
      bool is2Pole
    )
    {
      CreateAndPositionText(
        tr,
        "PANEL",
        "gmep",
        0.1872,
        0.75,
        0,
        layerName,
        new Point3d(startPoint.X + 0.231944251649111, startPoint.Y - 0.299822699224023, 0)
      );
      CreateAndPositionText(
        tr,
        "DESCRIPTION",
        "gmep",
        0.1248,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 0.305517965881791, startPoint.Y - 0.638118222684739, 0)
      );
      CreateAndPositionText(
        tr,
        "W",
        "gmep",
        0.101088,
        0.75,
        0,
        layerName,
        new Point3d(startPoint.X + 8.64365164909793, startPoint.Y - 0.145688865359394, 0)
      );
      CreateAndPositionText(
        tr,
        "VOLT AMPS",
        "gmep",
        0.11232,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 1.9015733562577, startPoint.Y - 0.532524377875689, 0)
      );
      CreateAndPositionText(
        tr,
        "L",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 2.97993751651882, startPoint.Y - 0.483601235896458, 0)
      );
      CreateAndPositionText(
        tr,
        "T",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 2.97993751651882, startPoint.Y - 0.59526740969153, 0)
      );
      CreateAndPositionText(
        tr,
        "G",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 2.97993751651882, startPoint.Y - 0.702157646684782, 0)
      );
      CreateAndPositionText(
        tr,
        "R",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 3.20889406685785, startPoint.Y - 0.482921120531671, 0)
      );
      CreateAndPositionText(
        tr,
        "E",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 3.20889406685785, startPoint.Y - 0.594587294326715, 0)
      );
      CreateAndPositionText(
        tr,
        "C",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 3.20889406685785, startPoint.Y - 0.701477531319966, 0)
      );
      CreateAndPositionText(
        tr,
        "M",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 3.43493724520761, startPoint.Y - 0.482921120531671, 0)
      );
      CreateAndPositionText(
        tr,
        "I",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 3.4427934214732, startPoint.Y - 0.594587294326715, 0)
      );
      CreateAndPositionText(
        tr,
        "S",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 3.43493724520761, startPoint.Y - 0.701477531319966, 0)
      );
      CreateAndPositionText(
        tr,
        "BKR",
        "gmep",
        0.09152,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 3.63691080609988, startPoint.Y - 0.61662650707666, 0)
      );
      CreateAndPositionText(
        tr,
        "CKT",
        "gmep",
        0.0832,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 3.94429929014041, startPoint.Y - 0.529332995532684, 0)
      );
      CreateAndPositionText(
        tr,
        " NO",
        "gmep",
        0.0832,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 3.90688892697108, startPoint.Y - 0.673306258645766, 0)
      );
      CreateAndPositionText(
        tr,
        "BUS",
        "gmep",
        0.11232,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 4.32282163085404, startPoint.Y - 0.527068325709052, 0)
      );
      CreateAndPositionText(
        tr,
        "CKT",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 4.88897460099258, startPoint.Y - 0.535275052777223, 0)
      );
      CreateAndPositionText(
        tr,
        " NO",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 4.85530527414039, startPoint.Y - 0.664850989579008, 0)
      );
      CreateAndPositionText(
        tr,
        "BKR",
        "gmep",
        0.082368,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 5.14497871612878, startPoint.Y - 0.612478980835647, 0)
      );
      CreateAndPositionText(
        tr,
        "M",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 5.4736003885796, startPoint.Y - 0.483601235896458, 0)
      );
      CreateAndPositionText(
        tr,
        "I",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 5.48257887574016, startPoint.Y - 0.59526740969153, 0)
      );
      CreateAndPositionText(
        tr,
        "S",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 5.4736003885796, startPoint.Y - 0.702157646684782, 0)
      );
      CreateAndPositionText(
        tr,
        "R",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 5.70588736710022, startPoint.Y - 0.482921120531671, 0)
      );
      CreateAndPositionText(
        tr,
        "E",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 5.70588736710022, startPoint.Y - 0.594587294326715, 0)
      );
      CreateAndPositionText(
        tr,
        "C",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 5.70588736710022, startPoint.Y - 0.701477531319966, 0)
      );
      CreateAndPositionText(
        tr,
        "L",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 5.93367350805136, startPoint.Y - 0.484281352862808, 0)
      );
      CreateAndPositionText(
        tr,
        "T",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 5.93367350805136, startPoint.Y - 0.595947526657881, 0)
      );
      CreateAndPositionText(
        tr,
        "G",
        "gmep",
        0.07488,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 5.93367350805136, startPoint.Y - 0.702837763651132, 0)
      );
      CreateAndPositionText(
        tr,
        "VOLT AMPS",
        "gmep",
        0.11232,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 6.32453930091015, startPoint.Y - 0.532297673821773, 0)
      );
      CreateAndPositionText(
        tr,
        "DESCRIPTION",
        "gmep",
        0.1248,
        0.75,
        256,
        "0",
        new Point3d(startPoint.X + 7.68034755863846, startPoint.Y - 0.636791573573134, 0)
      );
      CreateAndPositionText(
        tr,
        "LOCATION",
        "gmep",
        0.11232,
        0.75,
        0,
        layerName,
        new Point3d(startPoint.X + 2.32067207718262, startPoint.Y - 0.155059196495415, 0)
      );
      CreateAndPositionText(
        tr,
        "MAIN (AMP)",
        "gmep",
        0.11232,
        0.75,
        0,
        layerName,
        new Point3d(startPoint.X + 2.32089885857886, startPoint.Y - 0.338479316609039, 0)
      );
      CreateAndPositionText(
        tr,
        "BUS RATING",
        "gmep",
        0.11232,
        0.75,
        0,
        layerName,
        new Point3d(startPoint.X + 5.18507633525223, startPoint.Y - 0.155059196495415, 0)
      );
      CreateAndPositionText(
        tr,
        "FED FROM",
        "gmep",
        0.11232,
        0.75,
        0,
        layerName,
        new Point3d(startPoint.X + 5.18507633525223, startPoint.Y - 0.338479316609039, 0)
      );
      CreateAndPositionText(
        tr,
        "MOUNTING:",
        "gmep",
        0.11232,
        0.75,
        0,
        layerName,
        new Point3d(startPoint.X + 7.01560982102967, startPoint.Y - 0.329154148660905, 0)
      );
      CreateAndPositionText(
        tr,
        "V",
        "gmep",
        0.09984,
        0.75,
        0,
        layerName,
        new Point3d(startPoint.X + 7.60112268015148, startPoint.Y - 0.148231303238949, 0)
      );
      CreateAndPositionText(
        tr,
        "\u03A6",
        "gmep",
        0.101088,
        0.75,
        0,
        layerName,
        new Point3d(startPoint.X + 8.25325740318381, startPoint.Y - 0.151432601608803, 0)
      );

      if (is2Pole)
      {
        CreateAndPositionText(
          tr,
          "\u03A6" + "A",
          "gmep",
          0.1248,
          0.75,
          256,
          "0",
          new Point3d(startPoint.X + 1.87939466183889, startPoint.Y - 0.720370467604425, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "B",
          "gmep",
          0.1248,
          0.75,
          256,
          "0",
          new Point3d(startPoint.X + 2.50641160863188, startPoint.Y - 0.720370467604425, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "A",
          "gmep",
          0.1248,
          0.75,
          256,
          "0",
          new Point3d(startPoint.X + 4.19245469916268, startPoint.Y - 0.720370467604425, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "B",
          "gmep",
          0.1248,
          0.75,
          256,
          "0",
          new Point3d(startPoint.X + 4.59766144739842, startPoint.Y - 0.720370467604425, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "A",
          "gmep",
          0.1248,
          0.75,
          256,
          "0",
          new Point3d(startPoint.X + 6.2528343633212, startPoint.Y - 0.720370467604425, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "B",
          "gmep",
          0.1248,
          0.75,
          256,
          "0",
          new Point3d(startPoint.X + 6.91903366501083, startPoint.Y - 0.720370467604425, 0)
        );
      }
      else
      {
        CreateAndPositionText(
          tr,
          "\u03A6" + "A",
          "gmep",
          0.11232,
          0.75,
          256,
          "0",
          new Point3d(startPoint.X + 1.75923320841673, startPoint.Y - 0.715823582939777, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "B",
          "gmep",
          0.11232,
          0.75,
          256,
          "0",
          new Point3d(startPoint.X + 2.17200390182074, startPoint.Y - 0.714690056264089, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "C",
          "gmep",
          0.11232,
          0.75,
          256,
          "0",
          new Point3d(startPoint.X + 2.57149885762158, startPoint.Y - 0.718725420176355, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "A",
          "gmep",
          0.11232,
          0.75,
          256,
          "0",
          new Point3d(startPoint.X + 4.22229040316483, startPoint.Y - 0.714236644953189, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "B",
          "gmep",
          0.11232,
          0.75,
          256,
          "0",
          new Point3d(startPoint.X + 4.42655098606872, startPoint.Y - 0.714236644953189, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "C",
          "gmep",
          0.11232,
          0.75,
          256,
          "0",
          new Point3d(startPoint.X + 4.63417850165774, startPoint.Y - 0.713042660752734, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "A",
          "gmep",
          0.11232,
          0.75,
          256,
          "0",
          new Point3d(startPoint.X + 6.22324655852697, startPoint.Y - 0.71537017323044, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "B",
          "gmep",
          0.11232,
          0.75,
          256,
          "0",
          new Point3d(startPoint.X + 6.63397621936463, startPoint.Y - 0.714690057865624, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "C",
          "gmep",
          0.11232,
          0.75,
          256,
          "0",
          new Point3d(startPoint.X + 7.03324439586629, startPoint.Y - 0.718272010467018, 0)
        );
      }
    }

    private void CreateTextsWithPanelData(
      Transaction tr,
      string layerName,
      Point3d startPoint,
      Dictionary<string, object> panelData
    )
    {
      CreateAndPositionText(
        tr,
        (panelData["panel"] as string).Trim(),
        "gmep",
        0.1872,
        0.75,
        2,
        layerName,
        new Point3d(startPoint.X + 1.17828457810867, startPoint.Y - 0.299822699224023, 0)
      );
      CreateAndPositionText(
        tr,
        panelData["location"] as string,
        "gmep",
        0.09375,
        1,
        2,
        layerName,
        new Point3d(startPoint.X + 3.19605976175148, startPoint.Y - 0.137807184107345, 0)
      );
      CreateAndPositionText(
        tr,
        panelData["main"] as string,
        "gmep",
        0.09375,
        1,
        2,
        layerName,
        new Point3d(startPoint.X + 3.24033367283675, startPoint.Y - 0.32590837886957, 0)
      );
      CreateAndPositionText(
        tr,
        panelData["bus_rating"] as string,
        "gmep",
        0.09375,
        1,
        2,
        layerName,
        new Point3d(startPoint.X + 6.1073642121926, startPoint.Y - 0.137807184107345, 0)
      );
      CreateAndPositionText(
        tr,
        ("'" + (panelData["fed_from"] as string)).Replace("'PANEL ", "PANEL '") + "'",
        "gmep",
        0.09375,
        1,
        2,
        layerName,
        new Point3d(startPoint.X + 6.0073642121926, startPoint.Y - 0.32590837886957, 0)
      );
      CreateAndPositionText(
        tr,
        panelData["voltage1"] as string + "/" + panelData["voltage2"] as string,
        "gmep",
        0.09375,
        1,
        2,
        layerName,
        new Point3d(startPoint.X + 7.04393671550224, startPoint.Y - 0.141653203021775, 0)
      );
      CreateAndPositionText(
        tr,
        panelData["mounting"] as string,
        "gmep",
        0.09375,
        1,
        2,
        layerName,
        new Point3d(startPoint.X + 7.87802551675406, startPoint.Y - 0.331292901876935, 0)
      );
      CreateAndPositionText(
        tr,
        panelData["phase"] as string,
        "gmep",
        0.09375,
        1,
        2,
        layerName,
        new Point3d(startPoint.X + 8.1253996026328, startPoint.Y - 0.141653203021775, 0)
      );
      CreateAndPositionText(
        tr,
        panelData["wire"] as string,
        "gmep",
        0.09375,
        1,
        2,
        layerName,
        new Point3d(startPoint.X + 8.50104048135836, startPoint.Y - 0.141653203021775, 0)
      );
    }

    public void CreateLoadSummary(
      Dictionary<string, object> panelData,
      List<Dictionary<string, object>> panelStorage
    )
    {
      var (doc, db, ed) = PanelCommands.GetGlobals();
      var promptOptions = new PromptPointOptions("\nSelect top left corner point: ");
      var promptResult = ed.GetPoint(promptOptions);
      if (promptResult.Status != PromptStatus.OK)
        return;

      var spaceId =
        (db.TileMode == true)
          ? SymbolUtilityServices.GetBlockModelSpaceId(db)
          : SymbolUtilityServices.GetBlockPaperSpaceId(db);

      // Initial point
      var topRightCorner = promptResult.Value;

      double GetSafeDouble(object val)
      {
        if (!String.IsNullOrEmpty(val as string))
        {
          return Convert.ToDouble(val as string);
        }
        return 0;
      }

      bool GetSafeBoolean(string key)
      {
        if (panelData.TryGetValue(key, out object value))
        {
          if (value is bool boolValue)
          {
            return boolValue;
          }
          if (value is string stringValue)
          {
            return bool.TryParse(stringValue, out bool result) && result;
          }
        }
        return false;
      }

      string GetLoadName(string key)
      {
        if (Regex.IsMatch(key, @"^[a-fA-F0-9]{8}(-[a-fA-F0-9]{4}){3}-[a-fA-F0-9]{12}$"))
        {
          foreach (Dictionary<string, object> panel in panelStorage)
          {
            if ((panel["id"] as string).ToLower() == key.ToLower())
            {
              return "PANEL " + panel["panel"] as string;
            }
          }
        }
        return key;
      }

      double GetPanelKva(string key)
      {
        if (Regex.IsMatch(key, @"^[a-fA-F0-9]{8}(-[a-fA-F0-9]{4}){3}-[a-fA-F0-9]{12}$"))
        {
          foreach (Dictionary<string, object> panel in panelStorage)
          {
            if ((panel["id"] as string).ToLower() == key.ToLower())
            {
              return GetSafeDouble(panel["kva"] as string);
            }
          }
        }
        return 0;
      }

      using (var tr = db.TransactionManager.StartTransaction())
      {
        var btr = (BlockTableRecord)tr.GetObject(spaceId, OpenMode.ForWrite);

        var startPoint = new Point3d(topRightCorner.X, topRightCorner.Y, 0);

        List<string> descriptionLeft = (List<string>)panelData["description_left"];
        List<string> descriptionRight = (List<string>)panelData["description_right"];
        List<string> phaseALeft = (List<string>)panelData["phase_a_left"];
        List<string> phaseARight = (List<string>)panelData["phase_a_right"];
        List<string> phaseBLeft = (List<string>)panelData["phase_b_left"];
        List<string> phaseBRight = (List<string>)panelData["phase_b_right"];
        List<string> phaseCLeft = [];
        List<string> phaseCRight = [];
        int totalEntries = 0;
        for (int i = 0; i < descriptionLeft.Count; i++)
        {
          if (
            !String.IsNullOrEmpty(descriptionLeft[i])
            && descriptionLeft[i] != "SPARE"
            && descriptionLeft[i] != "SPACE"
          )
          {
            totalEntries++;
          }
          if (
            !String.IsNullOrEmpty(descriptionRight[i])
            && descriptionRight[i] != "SPARE"
            && descriptionRight[i] != "SPACE"
          )
          {
            totalEntries++;
          }
        }
        Table tb = new Table();
        List<double> kvaList = new List<double>();
        tb.TableStyle = db.Tablestyle;
        tb.Position = startPoint;
        tb.SetSize(totalEntries + 5, 3);
        tb.SetRowHeight(0.5);
        tb.Cells[0, 0].TextHeight = (0.125);
        tb.Cells[0, 0].TextString = $"{panelData["panel"] as string} LOAD SUMMARY";
        int tableRowIndex = 1;
        int increment = 2;
        tb.SetColumnWidth(2);
        if (panelData.ContainsKey("phase_c_left"))
        {
          increment = 3;
          phaseCLeft = (List<string>)panelData["phase_c_left"];
          phaseCRight = (List<string>)panelData["phase_c_right"];
        }
        for (int i = 0; i < descriptionLeft.Count; i += increment * 2)
        {
          if (
            !String.IsNullOrEmpty(descriptionLeft[i])
            && descriptionLeft[i] != "SPARE"
            && descriptionLeft[i] != "SPACE"
          )
          {
            double panelKva = GetPanelKva(descriptionLeft[i]);
            tb.Cells[tableRowIndex, 2].TextString = panelKva.ToString() + " KVA";
            kvaList.Add(panelKva);
            if (panelKva == 0)
            {
              double phA = GetSafeDouble(phaseALeft[i]);
              double phB = GetSafeDouble(phaseBLeft[i + 2]);
              double phC = 0;
              if (increment == 3)
              {
                phC = GetSafeDouble(phaseCLeft[i + 4]);
              }
              tb.Cells[tableRowIndex, 2].TextString =
                Math.Round((phA + phB + phC) / 1000, 1).ToString() + " KVA";
              kvaList.Add(Math.Round((phA + phB + phC) / 1000, 1));
            }
            tb.Cells[tableRowIndex, 0].TextString = $"{tableRowIndex}.";
            tb.Cells[tableRowIndex, 1].TextString = GetLoadName(descriptionLeft[i]);
            tableRowIndex++;
          }
          if (
            !String.IsNullOrEmpty(descriptionRight[i])
            && descriptionRight[i] != "SPARE"
            && descriptionRight[i] != "SPACE"
          )
          {
            double panelKva = GetPanelKva(descriptionRight[i]);
            tb.Cells[tableRowIndex, 2].TextString = panelKva.ToString() + " KVA";
            kvaList.Add(panelKva);
            if (panelKva == 0)
            {
              double phA = GetSafeDouble(phaseARight[i]);
              double phB = GetSafeDouble(phaseBRight[i + 2]);
              double phC = 0;
              if (increment == 3)
              {
                phC = GetSafeDouble(phaseCRight[i + 4]);
              }
              tb.Cells[tableRowIndex, 2].TextString =
                Math.Round((phA + phB + phC) / 1000, 1).ToString() + " KVA";
              kvaList.Add(Math.Round((phA + phB + phC) / 1000, 1));
            }
            tb.Cells[tableRowIndex, 0].TextString = $"{tableRowIndex}.";
            tb.Cells[tableRowIndex, 1].TextString = GetLoadName(descriptionRight[i]);
            tableRowIndex++;
          }
        }

        var textStyleId = GetTextStyleId("gmep");
        tb.Layer = "E-TXT1";
        for (int i = 0; i < totalEntries + 5; i++)
        {
          for (int j = 0; j < 3; j++)
          {
            tb.Cells[i, j].TextHeight = 0.1;
            tb.Cells[i, j].Alignment = CellAlignment.MiddleCenter;
            tb.Cells[i, j].TextStyleId = textStyleId;
          }
        }

        tb.Columns[0].Width = 0.5;
        for (int i = totalEntries + 1; i < totalEntries + 5; i++)
        {
          CellRange range = CellRange.Create(tb, i, 0, i, 1);
          tb.MergeCells(range);
        }
        tb.Cells[totalEntries + 1, 0].TextString = "TOTAL KVA";

        double kva = kvaList.Sum();
        tb.Cells[totalEntries + 1, 2].TextString = kva.ToString();

        bool usingSafetyFactor = GetSafeBoolean("using_safety_factor");
        double phaseVoltage = GetSafeDouble(panelData["voltage1"]);
        double lineVoltage = GetSafeDouble(panelData["voltage2"]);
        double yFactor = 1;
        int phase = 1;
        int wire = 3;
        if (increment == 3 && lineVoltage != 240)
        {
          yFactor = 1.732;
          phase = 3;
          wire = 4;
        }
        double feederAmps = GetSafeDouble(panelData["feeder_amps"]);

        double busSize = Convert.ToDouble(
          panelData.TryGetValue("bus_rating", out object bus)
            ? bus?.ToString().Replace("A", "") ?? "1"
            : "1"
        );

        double totalAmperage = 0;
        if (usingSafetyFactor)
        {
          double safetyFactor = Convert.ToDouble(
            panelData.TryGetValue("safety_factor", out object value)
              ? value?.ToString() ?? "1"
              : "1"
          );
          totalAmperage = Math.Round(kva * 1000 / lineVoltage / yFactor * safetyFactor, 1);
          if (safetyFactor == 0)
            safetyFactor = 1;
          tb.Cells[totalEntries + 2, 0].TextString = $"TOTAL KVA x{safetyFactor}";
          tb.Cells[totalEntries + 2, 2].TextString = Math.Round(kva * safetyFactor, 1).ToString();
          tb.Cells[totalEntries + 3, 0].TextString =
            $"TOTAL AMP @{phaseVoltage}/{lineVoltage}V-{phase}\u03A6-{wire}W";
          tb.Cells[totalEntries + 3, 2].TextString = totalAmperage.ToString();
          if (totalAmperage < busSize)
          {
            tb.Cells[totalEntries + 4, 0].TextString = "CONCLUSION";
            tb.Cells[totalEntries + 4, 2].TextString =
              $"{busSize}A SERVICE CAN HANDLE {totalAmperage}A LOAD.";
          }
          else
          {
            tb.Cells[totalEntries + 4, 0].TextString = "CONCLUSION";
            tb.Cells[totalEntries + 4, 2].TextString =
              $"{busSize}A SERVICE CANNOT HANDLE {totalAmperage}A LOAD.";
          }
        }
        else
        {
          totalAmperage = Math.Round(kva * 1000 / lineVoltage / yFactor, 1);
          tb.Cells[totalEntries + 2, 0].TextString =
            $"TOTAL AMP @{phaseVoltage}/{lineVoltage}V-{phase}\u03A6-{wire}W";
          tb.Cells[totalEntries + 2, 2].TextString = Math.Round(
              kva * 1000 / lineVoltage / yFactor,
              1
            )
            .ToString();
          if (totalAmperage < busSize)
          {
            tb.Cells[totalEntries + 3, 0].TextString = "CONCLUSION";
            tb.Cells[totalEntries + 3, 2].TextString =
              $"{busSize}A SERVICE CAN HANDLE {totalAmperage}A LOAD.";
          }
          else
          {
            tb.Cells[totalEntries + 3, 0].TextString = "CONCLUSION";
            tb.Cells[totalEntries + 3, 2].TextString =
              $"{busSize}A SERVICE CANNOT HANDLE {totalAmperage}A LOAD.";
          }
          tb.DeleteRows(totalEntries + 4, 1);
        }

        BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
        btr.AppendEntity(tb);
        tr.AddNewlyCreatedDBObject(tb, true);
        tr.Commit();
      }
    }

    public void CreatePanels(List<Dictionary<string, object>> panelDataList)
    {
      var (doc, db, ed) = PanelCommands.GetGlobals();

      if (panelDataList == null)
      {
        panelDataList = ImportExcelData(); // If not provided, import from Excel
      }

      var spaceId =
        (db.TileMode == true)
          ? SymbolUtilityServices.GetBlockModelSpaceId(db)
          : SymbolUtilityServices.GetBlockPaperSpaceId(db);

      // Get the insertion point from the user
      var promptOptions = new PromptPointOptions("\nSelect top right corner point: ");
      var promptResult = ed.GetPoint(promptOptions);
      if (promptResult.Status != PromptStatus.OK)
        return;

      // Initial point
      var topRightCorner = promptResult.Value;
      var originalTopRightCorner = promptResult.Value;

      // Lowest Y point
      double lowestY = topRightCorner.Y;

      var totalLevel = 0;
      var decreaseY = 0.0;

      int counter = 0;
      CREATEBLOCK();
      foreach (var panelData in panelDataList)
      {
        circuitNumLeft = 1;
        circuitNumRight = 2;
        if (panelData.TryGetValue("distribution_section", out object value))
        {
          if (value is bool boolValue && boolValue == false)
          {
            bool is2Pole = !panelData.ContainsKey("phase_c_left");
            var endPoint = new Point3d(0, 0, 0);

            using (var tr = db.TransactionManager.StartTransaction())
            {
              var btr = (BlockTableRecord)tr.GetObject(spaceId, OpenMode.ForWrite);

              // Create initial values
              var startPoint = new Point3d(topRightCorner.X - 8.9856, topRightCorner.Y, 0);
              var layerName = "0";

              // Create the independent header text objects
              CreateTextsWithoutPanelData(tr, layerName, startPoint, is2Pole);

              // Create the dependent header text objects
              CreateTextsWithPanelData(tr, layerName, startPoint, panelData);

              // Create breaker text objects
              totalLevel = ProcessTextData(tr, btr, startPoint, panelData, is2Pole);

              // Get end of data
              var endOfDataY = GetEndOfDataY(
                (List<string>)panelData["description_left"],
                startPoint
              );
              endPoint = new Point3d(topRightCorner.X, endOfDataY - 0.2533, 0);

              // Create all the data lines
              ProcessLineData(tr, btr, startPoint, endPoint, endOfDataY, is2Pole);

              // Create footer text objects
              CreateFooterText(tr, endPoint, panelData, is2Pole);

              // Create the middle lines
              CreateCenterLines(btr, tr, startPoint, endPoint, is2Pole);

              // Create the notes section
              if (panelData.ContainsKey("notes"))
              {
                var decrease = CreateNotes(
                  btr,
                  tr,
                  startPoint,
                  endPoint,
                  panelData["existing"] as string,
                  panelData["custom_title"] as string,
                  panelData["notes"] as List<string>
                );
                if (decrease > decreaseY)
                {
                  decreaseY = decrease;
                }
              }
              else
              {
                var decrease = CreateNotes(
                  btr,
                  tr,
                  startPoint,
                  endPoint,
                  panelData["existing"] as string,
                  null,
                  null
                );
                if (decrease > decreaseY)
                {
                  decreaseY = decrease;
                }
              }

              // Create the calculations section
              CreateCalculations(btr, tr, startPoint, endPoint, panelData);

              // Create the border of the panel
              CreateRectangle(btr, tr, topRightCorner, startPoint, endPoint, layerName);

              tr.Commit();
            }

            // Check if the endPoint.Y is the lowest point
            if (endPoint.Y < lowestY)
            {
              lowestY = endPoint.Y;
            }

            counter++;

            // After printing 3 panels, reset X and decrease Y by 5
            if (counter % 3 == 0)
            {
              topRightCorner = new Point3d(originalTopRightCorner.X, lowestY - 1.5 - decreaseY, 0);
              // Reset lowestY
              lowestY = topRightCorner.Y;
            }
            else
            {
              // Increase x-coordinate by 10 for the next panel
              topRightCorner = new Point3d(
                topRightCorner.X - (9.6 + (0.2 * totalLevel)),
                topRightCorner.Y,
                0
              );
            }
          }
        }
      }
    }

    private void SaveDataInJsonFileOnDesktop(object allXrefFileNames, string v)
    {
      string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
      string filePath = Path.Combine(desktopPath, v);

      string json = Newtonsoft.Json.JsonConvert.SerializeObject(
        allXrefFileNames,
        Newtonsoft.Json.Formatting.Indented
      );
      File.WriteAllText(filePath, json);
    }

    private static List<Dictionary<string, object>> ImportExcelData()
    {
      var (doc, db, ed) = GetGlobals();

      var openFileDialog = new System.Windows.Forms.OpenFileDialog
      {
        Filter = "Excel Files|*.xlsx;*.xls",
        Title = "Select Excel File",
      };

      List<Dictionary<string, object>> panels = new List<Dictionary<string, object>>();

      if (openFileDialog.ShowDialog() == DialogResult.OK)
      {
        string filePath = openFileDialog.FileName;
        try
        {
          FileInfo fileInfo = GetFileInfo(filePath);
          using (var package = GetExcelPackage(fileInfo))
          {
            ExcelWorkbook workbook = package.Workbook;
            if (ValidateWorkbook(workbook))
            {
              foreach (var selectedWorksheet in workbook.Worksheets)
              {
                if (selectedWorksheet.Name.ToLower().Contains("panel"))
                {
                  panels.AddRange(ProcessWorksheet(selectedWorksheet));
                }
              }
            }
          }
        }
        catch (FileNotFoundException ex)
        {
          HandleExceptions(ex);
        }
        catch (Autodesk.AutoCAD.Runtime.Exception ex)
        {
          HandleExceptions(ex);
        }
      }
      else
      {
        Console.WriteLine("No file selected.");
      }

      return panels;
    }

    public static (Document doc, Database db, Editor ed) GetGlobals()
    {
      var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      var db = doc.Database;
      var ed = doc.Editor;

      return (doc, db, ed);
    }

    public void CreatePanel(Dictionary<string, object> panelData)
    {
      List<Dictionary<string, object>> panels = new List<Dictionary<string, object>>();
      panels.Add(panelData);
      CreatePanels(panels);
    }

    private static FileInfo GetFileInfo(string filePath)
    {
      var fileInfo = new FileInfo(filePath);
      if (!fileInfo.Exists)
      {
        throw new FileNotFoundException($"The file {filePath} does not exist.");
      }
      return fileInfo;
    }

    private static ExcelPackage GetExcelPackage(FileInfo fileInfo)
    {
      return new ExcelPackage(fileInfo);
    }

    private static bool ValidateWorkbook(ExcelWorkbook workbook)
    {
      if (workbook == null)
      {
        Console.WriteLine("Workbook not found.");
        return false;
      }
      return true;
    }

    private static Dictionary<string, object> ProcessThreePolePanel(
      ExcelWorksheet selectedWorksheet,
      int row,
      int col
    )
    {
      List<string> descriptionLeft = new List<string>();
      List<string> phaseALeft = new List<string>();
      List<string> phaseBLeft = new List<string>();
      List<string> phaseCLeft = new List<string>();
      List<string> breakerLeft = new List<string>();
      List<string> circuitLeft = new List<string>();
      List<string> circuitRight = new List<string>();
      List<string> breakerRight = new List<string>();
      List<string> phaseARight = new List<string>();
      List<string> phaseBRight = new List<string>();
      List<string> phaseCRight = new List<string>();
      List<string> descriptionRight = new List<string>();

      int panelRow = row + 4;
      int lastRow = panelRow;

      while (selectedWorksheet.Cells[lastRow + 2, col + 6].Value != null)
      {
        lastRow += 2;
      }

      lastRow += 1;

      List<bool> descriptionLeftHighlights = new List<bool>();
      List<bool> descriptionRightHighlights = new List<bool>();
      List<bool> breakerLeftHighlights = new List<bool>();
      List<bool> breakerRightHighlights = new List<bool>();

      for (int i = panelRow; i <= lastRow; i++)
      {
        string description = "";
        if (
          selectedWorksheet.Cells[i, col].Value == null
          && selectedWorksheet.Cells[i, col + 5].Value != null
        )
        {
          description = "SPARE";
        }
        else
        {
          description = selectedWorksheet.Cells[i, col].Value?.ToString().ToUpper() ?? "SPACE";
        }
        string descriptionR = "";
        if (
          selectedWorksheet.Cells[i, col + 12].Value == null
          && selectedWorksheet.Cells[i, col + 8].Value != null
        )
        {
          descriptionR = "SPARE";
        }
        else
        {
          descriptionR =
            selectedWorksheet.Cells[i, col + 12].Value?.ToString().ToUpper() ?? "SPACE";
        }
        string phaseA = selectedWorksheet.Cells[i, col + 2].Value?.ToString() ?? "0";
        string phaseB = selectedWorksheet.Cells[i, col + 3].Value?.ToString() ?? "0";
        string phaseC = selectedWorksheet.Cells[i, col + 4].Value?.ToString() ?? "0";
        string breaker = selectedWorksheet.Cells[i, col + 5].Value?.ToString() ?? "";
        string circuitL = selectedWorksheet.Cells[i, col + 6].Value?.ToString() ?? "";
        string circuitR = selectedWorksheet.Cells[i, col + 7].Value?.ToString() ?? "";
        string breakerR = selectedWorksheet.Cells[i, col + 8].Value?.ToString() ?? "";
        string phaseAR = selectedWorksheet.Cells[i, col + 9].Value?.ToString() ?? "0";
        string phaseBR = selectedWorksheet.Cells[i, col + 10].Value?.ToString() ?? "0";
        string phaseCR = selectedWorksheet.Cells[i, col + 11].Value?.ToString() ?? "0";

        bool isLeftHighlighted =
          selectedWorksheet.Cells[i, col].Style.Fill.BackgroundColor.LookupColor() != "#FF000000";
        bool isRightHighlighted =
          selectedWorksheet.Cells[i, col + 12].Style.Fill.BackgroundColor.LookupColor()
          != "#FF000000";
        bool isLeftBreakerHighlighted =
          selectedWorksheet.Cells[i, col + 5].Style.Fill.BackgroundColor.LookupColor()
          != "#FF000000";
        bool isRightBreakerHighlighted =
          selectedWorksheet.Cells[i, col + 8].Style.Fill.BackgroundColor.LookupColor()
          != "#FF000000";

        descriptionLeft.Add(description);
        phaseALeft.Add(phaseA);
        phaseBLeft.Add(phaseB);
        phaseCLeft.Add(phaseC);
        breakerLeft.Add(breaker);
        circuitLeft.Add(circuitL);
        circuitRight.Add(circuitR);
        breakerRight.Add(breakerR);
        phaseARight.Add(phaseAR);
        phaseBRight.Add(phaseBR);
        phaseCRight.Add(phaseCR);
        descriptionRight.Add(descriptionR);
        descriptionLeftHighlights.Add(isLeftHighlighted);
        descriptionRightHighlights.Add(isRightHighlighted);
        breakerLeftHighlights.Add(isLeftBreakerHighlighted);
        breakerRightHighlights.Add(isRightBreakerHighlighted);
      }

      string panelCellValue = selectedWorksheet.Cells[row, col + 2].Value?.ToString() ?? "";
      if (!panelCellValue.StartsWith("'") || !panelCellValue.EndsWith("'"))
      {
        panelCellValue = "'" + panelCellValue.Trim('\'') + "'";
      }

      var panel = new Dictionary<string, object>
      {
        ["panel"] = panelCellValue,
        ["location"] = selectedWorksheet.Cells[row, col + 5].Value?.ToString() ?? "",
        ["bus_rating"] = selectedWorksheet.Cells[row, col + 9].Value?.ToString() ?? "",
        ["voltage1"] = selectedWorksheet.Cells[row, col + 10].Value?.ToString() ?? "0",
        ["voltage2"] = selectedWorksheet.Cells[row, col + 11].Value?.ToString() ?? "0",
        ["phase"] = selectedWorksheet.Cells[row, col + 12].Value?.ToString() ?? "",
        ["wire"] = selectedWorksheet.Cells[row, col + 13].Value?.ToString() ?? "",
        ["main"] = selectedWorksheet.Cells[row + 1, col + 5].Value?.ToString() ?? "",
        ["mounting"] = selectedWorksheet.Cells[row + 1, col + 12].Value?.ToString() ?? "",
        ["subtotal_a"] = selectedWorksheet.Cells[row + 2, col + 17].Value?.ToString() ?? "0",
        ["subtotal_b"] = selectedWorksheet.Cells[row + 2, col + 18].Value?.ToString() ?? "0",
        ["subtotal_c"] = selectedWorksheet.Cells[row + 2, col + 19].Value?.ToString() ?? "0",
        ["total_va"] = selectedWorksheet.Cells[row + 4, col + 17].Value?.ToString() ?? "0",
        ["lcl"] = selectedWorksheet.Cells[row + 7, col + 17].Value?.ToString() ?? "0",
        ["lcl_125"] = "0",
        ["total_other_load"] = selectedWorksheet.Cells[row + 10, col + 17].Value?.ToString() ?? "0",
        ["kva"] = selectedWorksheet.Cells[row + 13, col + 17].Value?.ToString() ?? "0",
        ["feeder_amps"] = selectedWorksheet.Cells[row + 16, col + 17].Value?.ToString() ?? "0",
        ["existing"] = selectedWorksheet.Cells[row + 2, col + 20].Value?.ToString() ?? "",
        ["description_left_highlights"] = descriptionLeftHighlights,
        ["description_right_highlights"] = descriptionRightHighlights,
        ["breaker_left_highlights"] = breakerLeftHighlights,
        ["breaker_right_highlights"] = breakerRightHighlights,
        ["description_left"] = descriptionLeft,
        ["phase_a_left"] = phaseALeft,
        ["phase_b_left"] = phaseBLeft,
        ["phase_c_left"] = phaseCLeft,
        ["breaker_left"] = breakerLeft,
        ["circuit_left"] = circuitLeft,
        ["circuit_right"] = circuitRight,
        ["breaker_right"] = breakerRight,
        ["phase_a_right"] = phaseARight,
        ["phase_b_right"] = phaseBRight,
        ["phase_c_right"] = phaseCRight,
        ["description_right"] = descriptionRight,
      };

      ReplaceInPanel(panel, "voltage2", "V");
      ReplaceInPanel(panel, "phase", "PH");
      ReplaceInPanel(panel, "wire", "W");

      return panel;
    }

    private static Dictionary<string, object> ProcessTwoPolePanel(
      ExcelWorksheet selectedWorksheet,
      int row,
      int col
    )
    {
      List<string> descriptionLeft = new List<string>();
      List<string> phaseALeft = new List<string>();
      List<string> phaseBLeft = new List<string>();
      List<string> breakerLeft = new List<string>();
      List<string> circuitLeft = new List<string>();
      List<string> circuitRight = new List<string>();
      List<string> breakerRight = new List<string>();
      List<string> phaseARight = new List<string>();
      List<string> phaseBRight = new List<string>();
      List<string> descriptionRight = new List<string>();
      List<bool> descriptionLeftHighlights = new List<bool>();
      List<bool> descriptionRightHighlights = new List<bool>();
      List<bool> breakerLeftHighlights = new List<bool>();
      List<bool> breakerRightHighlights = new List<bool>();

      int panelRow = row + 4;
      int lastRow = panelRow;

      // check for a circuit column cell value of NULL to end the loop
      while (selectedWorksheet.Cells[lastRow + 2, col + 6].Value != null)
      {
        lastRow += 2;
      }

      lastRow += 1;

      // add cell values to lists
      for (int i = panelRow; i <= lastRow; i++)
      {
        string description = "";
        if (
          selectedWorksheet.Cells[i, col].Value == null
          && selectedWorksheet.Cells[i, col + 5].Value != null
        )
        {
          description = "SPARE";
        }
        else
        {
          description = selectedWorksheet.Cells[i, col].Value?.ToString().ToUpper() ?? "SPACE";
        }
        string descriptionR = "";
        if (
          selectedWorksheet.Cells[i, col + 11].Value == null
          && selectedWorksheet.Cells[i, col + 8].Value != null
        )
        {
          descriptionR = "SPARE";
        }
        else
        {
          descriptionR =
            selectedWorksheet.Cells[i, col + 11].Value?.ToString().ToUpper() ?? "SPACE";
        }
        string phaseA = selectedWorksheet.Cells[i, col + 3].Value?.ToString() ?? "0";
        string phaseB = selectedWorksheet.Cells[i, col + 4].Value?.ToString() ?? "0";
        string breakerL = selectedWorksheet.Cells[i, col + 5].Value?.ToString() ?? "";
        string circuitL = selectedWorksheet.Cells[i, col + 6].Value?.ToString() ?? "";
        string circuitR = selectedWorksheet.Cells[i, col + 7].Value?.ToString() ?? "";
        string breakerR = selectedWorksheet.Cells[i, col + 8].Value?.ToString() ?? "";
        string phaseAR = selectedWorksheet.Cells[i, col + 9].Value?.ToString() ?? "0";
        string phaseBR = selectedWorksheet.Cells[i, col + 10].Value?.ToString() ?? "0";

        bool isLeftHighlighted =
          selectedWorksheet.Cells[i, col].Style.Fill.BackgroundColor.LookupColor() != "#FF000000";
        bool isRightHighlighted =
          selectedWorksheet.Cells[i, col + 11].Style.Fill.BackgroundColor.LookupColor()
          != "#FF000000";
        bool isLeftBreakerHighlighted =
          selectedWorksheet.Cells[i, col + 5].Style.Fill.BackgroundColor.LookupColor()
          != "#FF000000";
        bool isRightBreakerHighlighted =
          selectedWorksheet.Cells[i, col + 8].Style.Fill.BackgroundColor.LookupColor()
          != "#FF000000";

        descriptionLeft.Add(description);
        phaseALeft.Add(phaseA);
        phaseBLeft.Add(phaseB);
        breakerLeft.Add(breakerL);
        circuitLeft.Add(circuitL);
        circuitRight.Add(circuitR);
        breakerRight.Add(breakerR);
        phaseARight.Add(phaseAR);
        phaseBRight.Add(phaseBR);
        descriptionRight.Add(descriptionR);
        descriptionLeftHighlights.Add(isLeftHighlighted);
        descriptionRightHighlights.Add(isRightHighlighted);
        breakerLeftHighlights.Add(isLeftBreakerHighlighted);
        breakerRightHighlights.Add(isRightBreakerHighlighted);
      }

      string panelCellValue = selectedWorksheet.Cells[row, col + 2].Value?.ToString() ?? "";
      if (!panelCellValue.StartsWith("'") || !panelCellValue.EndsWith("'"))
      {
        panelCellValue = "'" + panelCellValue.Trim('\'') + "'";
      }

      var panel = new Dictionary<string, object>
      {
        ["panel"] = panelCellValue,
        ["location"] = selectedWorksheet.Cells[row, col + 5].Value?.ToString() ?? "",
        ["bus_rating"] = selectedWorksheet.Cells[row, col + 9].Value?.ToString() ?? "",
        ["voltage1"] = selectedWorksheet.Cells[row, col + 10].Value?.ToString() ?? "0",
        ["voltage2"] = selectedWorksheet.Cells[row, col + 11].Value?.ToString() ?? "0",
        ["phase"] = selectedWorksheet.Cells[row, col + 12].Value?.ToString() ?? "",
        ["wire"] = selectedWorksheet.Cells[row, col + 13].Value?.ToString() ?? "",
        ["main"] = selectedWorksheet.Cells[row + 1, col + 5].Value?.ToString() ?? "",
        ["mounting"] = selectedWorksheet.Cells[row + 1, col + 12].Value?.ToString() ?? "",
        ["subtotal_a"] = selectedWorksheet.Cells[row + 2, col + 17].Value?.ToString() ?? "0",
        ["subtotal_b"] = selectedWorksheet.Cells[row + 2, col + 18].Value?.ToString() ?? "0",
        ["subtotal_c"] = selectedWorksheet.Cells[row + 2, col + 19].Value?.ToString() ?? "0",
        ["total_va"] = selectedWorksheet.Cells[row + 4, col + 17].Value?.ToString() ?? "0",
        ["lcl"] = selectedWorksheet.Cells[row + 7, col + 17].Value?.ToString() ?? "0",
        ["lcl_125"] = "0",
        ["total_other_load"] = selectedWorksheet.Cells[row + 10, col + 17].Value?.ToString() ?? "0",
        ["kva"] = selectedWorksheet.Cells[row + 13, col + 17].Value?.ToString() ?? "0",
        ["feeder_amps"] = selectedWorksheet.Cells[row + 16, col + 17].Value?.ToString() ?? "0",
        ["existing"] = selectedWorksheet.Cells[row + 2, col + 20].Value?.ToString() ?? "",
        ["description_left_highlights"] = descriptionLeftHighlights,
        ["description_right_highlights"] = descriptionRightHighlights,
        ["breaker_left_highlights"] = breakerLeftHighlights,
        ["breaker_right_highlights"] = breakerRightHighlights,
        ["description_left"] = descriptionLeft,
        ["phase_a_left"] = phaseALeft,
        ["phase_b_left"] = phaseBLeft,
        ["breaker_left"] = breakerLeft,
        ["circuit_left"] = circuitLeft,
        ["circuit_right"] = circuitRight,
        ["breaker_right"] = breakerRight,
        ["phase_a_right"] = phaseARight,
        ["phase_b_right"] = phaseBRight,
        ["description_right"] = descriptionRight,
      };

      ReplaceInPanel(panel, "voltage2", "V");
      ReplaceInPanel(panel, "phase", "PH");
      ReplaceInPanel(panel, "wire", "W");

      return panel;
    }

    private static void ReplaceInPanel(
      Dictionary<string, object> panel,
      string key,
      string toRemove
    )
    {
      if (panel.ContainsKey(key))
      {
        if (panel[key] is string value)
        {
          value = value.Replace(toRemove, "");
          panel[key] = value;
        }
      }
    }

    private static List<Dictionary<string, object>> ProcessWorksheet(
      ExcelWorksheet selectedWorksheet
    )
    {
      var panels = new List<Dictionary<string, object>>();
      int rowCount = selectedWorksheet.Dimension.Rows;
      int colCount = selectedWorksheet.Dimension.Columns;

      for (int row = 1; row <= rowCount; row++)
      {
        for (int col = 1; col <= colCount; col++)
        {
          string cellValue = selectedWorksheet.Cells[row, col].Value?.ToString();
          string phaseAMaybe = selectedWorksheet.Cells[row + 3, col + 2].Value?.ToString();
          if (cellValue == "PANEL:" && phaseAMaybe == "PH A")
          {
            panels.Add(ProcessThreePolePanel(selectedWorksheet, row, col));
          }
          else if (cellValue == "PANEL:")
          {
            panels.Add(ProcessTwoPolePanel(selectedWorksheet, row, col));
          }
        }
      }
      return panels;
    }

    private static void HandleExceptions(System.Exception ex)
    {
      Console.WriteLine(ex.Message);
    }

    public void KeepBreakersGivenPoints(
      Point3d point1,
      Point3d point2,
      Point3d point3,
      string content
    )
    {
      var (doc, db, ed) = PanelCommands.GetGlobals();

      using (var tr = db.TransactionManager.StartTransaction())
      {
        var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

        var activeSpaceId = db.CurrentSpaceId;
        var btr = (BlockTableRecord)tr.GetObject(activeSpaceId, OpenMode.ForWrite);

        if (point1.Y < point2.Y)
        {
          Point3d tempPoint = point1;
          point1 = point2;
          point2 = tempPoint;
        }

        var direction = point3.X > point1.X ? 1 : -1;
        var dist = (point1 - point2).Length;

        var line1Start = new Point3d(point1.X + direction * 0.05, point1.Y, 0);
        var line1End = new Point3d(line1Start.X + direction * 0.2, line1Start.Y, 0);
        var line2Start = new Point3d(line1Start.X, point2.Y, 0);
        var line2End = new Point3d(line1End.X, line2Start.Y, 0);

        string layerName = CreateOrGetLayer("E-TEXT", db, tr);

        var line1 = new Line(line1Start, line1End)
        {
          Layer = layerName,
          Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 2),
        };
        var line2 = new Line(line2Start, line2End)
        {
          Layer = layerName,
          Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 2),
        };

        var mid1 = new Point3d((line1Start.X + line1End.X) / 2, (line1Start.Y + line1End.Y) / 2, 0);
        var mid2 = new Point3d((line2Start.X + line2End.X) / 2, (line2Start.Y + line2End.Y) / 2, 0);
        var mid3 = new Point3d((mid1.X + mid2.X) / 2, (mid1.Y + mid2.Y) / 2, 0);

        var circleTop = new Point3d(mid3.X, mid3.Y + 0.09, 0);
        var circleBottom = new Point3d(mid3.X, mid3.Y - 0.09, 0);

        var line3 = new Line(mid1, circleTop)
        {
          Layer = layerName,
          Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 2),
        };
        var line4 = new Line(mid2, circleBottom)
        {
          Layer = layerName,
          Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 2),
        };

        if (dist > 0.3)
        {
          btr.AppendEntity(line1);
          btr.AppendEntity(line2);
          tr.AddNewlyCreatedDBObject(line1, true);
          tr.AddNewlyCreatedDBObject(line2, true);
        }

        CreateCircle(btr, tr, mid3, 0.09, 2, false);
        CreateCircleText(btr, tr, mid3, 0.09, 2, "gmep", content);

        if (dist > 0.3)
        {
          btr.AppendEntity(line3);
          btr.AppendEntity(line4);
          tr.AddNewlyCreatedDBObject(line3, true);
          tr.AddNewlyCreatedDBObject(line4, true);
        }

        tr.Commit();
      }
    }

    private void CreateCircleText(
      BlockTableRecord btr,
      Transaction tr,
      Point3d centerPoint,
      double height,
      int colorIndex,
      string textStyle,
      string content
    )
    {
      DBText text = new DBText();
      text.SetDatabaseDefaults();
      text.Height = height;
      text.TextString = content;
      text.HorizontalMode = TextHorizontalMode.TextCenter;
      text.VerticalMode = TextVerticalMode.TextVerticalMid;

      TextStyleTable textStyleTable = (TextStyleTable)
        tr.GetObject(btr.Database.TextStyleTableId, OpenMode.ForRead);

      if (textStyleTable.Has(textStyle))
      {
        text.TextStyleId = tr.GetObject(textStyleTable[textStyle], OpenMode.ForRead).ObjectId;
      }
      else
      {
        text.TextStyleId = tr.GetObject(textStyleTable["gmep"], OpenMode.ForRead).ObjectId;
      }

      text.ColorIndex = colorIndex;

      text.AlignmentPoint = centerPoint;

      btr.AppendEntity(text);
      tr.AddNewlyCreatedDBObject(text, true);
    }

    public string CreateOrGetLayer(string layerName, Database db, Transaction tr)
    {
      var lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

      if (!lt.Has(layerName)) // check if layer exists
      {
        lt.UpgradeOpen(); // switch to write mode
        LayerTableRecord ltr = new LayerTableRecord();
        ltr.Name = layerName;
        lt.Add(ltr);
        tr.AddNewlyCreatedDBObject(ltr, true);
      }

      return layerName;
    }

    private void CreateCalculations(
      BlockTableRecord btr,
      Transaction tr,
      Point3d startPoint,
      Point3d endPoint,
      Dictionary<string, object> panelData
    )
    {
      var (_, _, ed) = GetGlobals();

      // Helper function to safely get string value from dictionary
      string GetSafeString(string key)
      {
        return panelData.TryGetValue(key, out object value) ? value?.ToString() ?? "" : "0";
      }

      // Helper function to safely parse double
      bool TryParseDouble(string key, out double result)
      {
        result = 0;
        string value = GetSafeString(key);
        return !string.IsNullOrEmpty(value) && double.TryParse(value, out result);
      }

      // KVA calculation
      if (TryParseDouble("kva", out double kvaValue))
      {
        CreateAndPositionRightText(
          tr,
          Math.Round(kvaValue, 1).ToString("0.0") + " KVA",
          "gmep",
          0.09375,
          1,
          2,
          "PNLTXT",
          new Point3d(endPoint.X - 6.69695957617801, endPoint.Y - 0.785594790702817, 0)
        );
      }
      else
      {
        ed.WriteMessage($"Error: Unable to convert 'kva' value to double: {GetSafeString("kva")}");
      }

      // Feeder Amps calculation
      if (TryParseDouble("feeder_amps", out double feederAmpsValue))
      {
        CreateAndPositionRightText(
          tr,
          Math.Round(feederAmpsValue, 1).ToString("0.0") + " A",
          "gmep",
          0.09375,
          1,
          2,
          "PNLTXT",
          new Point3d(endPoint.X - 6.70142386189229, endPoint.Y - 0.970762733814496, 0)
        );
      }
      else
      {
        ed.WriteMessage(
          $"Error: Unable to convert 'feeder_amps' value to double: {GetSafeString("feeder_amps")}"
        );
      }

      // Create the calculation lines
      CreateLine(
        tr,
        btr,
        endPoint.X - 6.17759999999998,
        endPoint.Y - 0.0846396524177919,
        endPoint.X - 8.98559999999998,
        endPoint.Y - 0.0846396524177919,
        "0"
      );
      CreateLine(
        tr,
        btr,
        endPoint.X - 6.17759999999998,
        endPoint.Y - 1.02063965241777,
        endPoint.X - 6.17759999999998,
        endPoint.Y - 0.0846396524177919,
        "0"
      );
      CreateLine(
        tr,
        btr,
        endPoint.X - 8.98559999999998,
        endPoint.Y - 1.02063965241777,
        endPoint.X - 8.98559999999998,
        endPoint.Y - 0.0846396524177919,
        "0"
      );
      CreateLine(
        tr,
        btr,
        endPoint.X - 6.17759999999998,
        endPoint.Y - 1.02063965241777,
        endPoint.X - 8.98559999999998,
        endPoint.Y - 1.02063965241777,
        "0"
      );
      CreateLine(
        tr,
        btr,
        endPoint.X - 6.17759999999998,
        endPoint.Y - 0.833439652417809,
        endPoint.X - 8.98559999999998,
        endPoint.Y - 0.833439652417809,
        "0"
      );
      CreateLine(
        tr,
        btr,
        endPoint.X - 6.17759999999998,
        endPoint.Y - 0.64623965241779,
        endPoint.X - 8.98559999999998,
        endPoint.Y - 0.64623965241779,
        "0"
      );
      CreateLine(
        tr,
        btr,
        endPoint.X - 6.17759999999998,
        endPoint.Y - 0.459039652417772,
        endPoint.X - 8.98559999999998,
        endPoint.Y - 0.459039652417772,
        "0"
      );
      CreateLine(
        tr,
        btr,
        endPoint.X - 6.17759999999998,
        endPoint.Y - 0.27183965241781,
        endPoint.X - 8.98559999999998,
        endPoint.Y - 0.27183965241781,
        "0"
      );
      CreateLine(
        tr,
        btr,
        endPoint.X - 6.17759999999998,
        endPoint.Y - 0.0846396524177919,
        endPoint.X - 8.98559999999998,
        endPoint.Y - 0.0846396524177919,
        "0"
      );

      // Create the text
      CreateAndPositionText(
        tr,
        "TOTAL CONNECTED VA",
        "gmep",
        0.1248,
        0.75,
        256,
        "0",
        new Point3d(endPoint.X - 8.93821353998555, endPoint.Y - 0.244065644556514, 0)
      );
      CreateAndPositionRightText(
        tr,
        GetSafeString("total_va"),
        "gmep",
        0.09375,
        1,
        2,
        "PNLTXT",
        new Point3d(endPoint.X - 6.69695957617801, endPoint.Y - 0.222040136230106, 0)
      );
      CreateAndPositionText(
        tr,
        "=",
        "gmep",
        0.1248,
        0.75,
        256,
        "0",
        new Point3d(endPoint.X - 7.03028501835593, endPoint.Y - 0.242614932747216, 0)
      );
      CreateAndPositionText(
        tr,
        "LCL @ 125 %          ",
        "gmep",
        0.1248,
        0.75,
        256,
        "0",
        new Point3d(endPoint.X - 8.91077927366155, endPoint.Y - 0.432165907882307, 0)
      );
      CreateAndPositionRightText(
        tr,
        GetSafeString("lcl"),
        "gmep",
        0.09375,
        1,
        2,
        "PNLTXT",
        new Point3d(endPoint.X - 7.59414061117746, endPoint.Y - 0.413648726513742, 0)
      );
      CreateAndPositionText(
        tr,
        "=",
        "gmep",
        0.1248,
        0.75,
        256,
        "0",
        new Point3d(endPoint.X - 7.03028501835593, endPoint.Y - 0.437756414851634, 0)
      );
      CreateAndPositionRightText(
        tr,
        GetSafeString("lcl125"),
        "gmep",
        0.09375,
        1,
        2,
        "PNLTXT",
        new Point3d(endPoint.X - 6.69695957617801, endPoint.Y - 0.413648726513742, 0)
      );
      CreateAndPositionText(
        tr,
        "LML @ 125 %",
        "gmep",
        0.1248,
        0.75,
        256,
        "0",
        new Point3d(endPoint.X - 8.91077927366155, endPoint.Y - 0.616854044919108, 0)
      );
      CreateAndPositionRightText(
        tr,
        GetSafeString("lml"),
        "gmep",
        0.09375,
        1,
        2,
        "PNLTXT",
        new Point3d(endPoint.X - 7.59414061117746, endPoint.Y - 0.618180694030713, 0)
      );
      CreateAndPositionText(
        tr,
        "=",
        "gmep",
        0.1248,
        0.75,
        256,
        "0",
        new Point3d(endPoint.X - 7.03028501835593, endPoint.Y - 0.618180694030713, 0)
      );
      CreateAndPositionRightText(
        tr,
        GetSafeString("lml125"),
        "gmep",
        0.09375,
        1,
        2,
        "PNLTXT",
        new Point3d(endPoint.X - 6.69695957617801, endPoint.Y - 0.618180694030713, 0)
      );
      CreateAndPositionText(
        tr,
        "PANEL LOAD",
        "gmep",
        0.1248,
        0.75,
        256,
        "0",
        new Point3d(endPoint.X - 8.92075537050664, endPoint.Y - 0.804954308244959, 0)
      );
      CreateAndPositionText(
        tr,
        "=",
        "gmep",
        0.1248,
        0.75,
        256,
        "0",
        new Point3d(endPoint.X - 7.03028501835593, endPoint.Y - 0.809218166102625, 0)
      );
      CreateAndPositionText(
        tr,
        "FEEDER AMPS",
        "gmep",
        0.1248,
        0.75,
        256,
        "0",
        new Point3d(endPoint.X - 8.9120262857673, endPoint.Y - 0.994381220682413, 0)
      );
      CreateAndPositionText(
        tr,
        "=",
        "gmep",
        0.1248,
        0.75,
        256,
        "0",
        new Point3d(endPoint.X - 7.03028501835593, endPoint.Y - 0.998928989062989, 0)
      );
    }

    private double CreateNotes(
      BlockTableRecord btr,
      Transaction tr,
      Point3d startPoint,
      Point3d endPoint,
      string panelType,
      string customTitle,
      List<string> customNotes
    )
    {
      string title;
      if (!string.IsNullOrEmpty(customTitle))
      {
        title = customTitle;
      }
      else
      {
        switch (panelType.ToLower())
        {
          case "existing":
            title = "(EXISTING PANEL)";
            break;

          case "relocated":
            title = "(EXISTING TO BE RELOCATED PANEL)";
            break;

          default:
            title = "(NEW PANEL)";
            break;
        }
      }

      // Create the horizontal lines
      double y_initial = 0.0846396524177919;
      double y_increment = 0.1872;
      int lines = 5;
      int lines_of_text = 0;
      double yOffset = y_initial;

      double decrease_y = 0;

      CreateAndPositionText(
        tr,
        title,
        "gmep",
        0.1498,
        0.75,
        2,
        "0",
        new Point3d(startPoint.X + 0.236635303895696, startPoint.Y + 0.113254677317428, 0)
      );
      CreateAndPositionText(
        tr,
        "NOTES:",
        "gmep",
        0.1248,
        0.75,
        256,
        "0",
        new Point3d(endPoint.X - 5.96783070435049, endPoint.Y - 0.23875904811004, 0)
      );

      if (customNotes == null)
      {
        if (panelType.ToLower() == "existing" || panelType.ToLower() == "relocated")
        {
          CreateAndPositionText(
            tr,
            "DENOTES EXISTING CIRCUIT BREAKER TO REMAIN; ALL OTHERS ARE NEW",
            "gmep",
            0.09375,
            1,
            2,
            "0",
            new Point3d(endPoint.X - 5.61904201783966, endPoint.Y - 0.405747901076808, 0)
          );
          CreateAndPositionText(
            tr,
            "TO MATCH EXISTING.",
            "gmep",
            0.09375,
            1,
            2,
            "0",
            new Point3d(endPoint.X - 5.61904201783966, endPoint.Y - 0.610352149436778, 0)
          );
        }
        else
        {
          CreateAndPositionText(
            tr,
            "65 KAIC SERIES RATED OR MATCH FAULT CURRENT AT SITE.",
            "gmep",
            0.09375,
            1,
            2,
            "0",
            new Point3d(endPoint.X - 5.61904201783966, endPoint.Y - 0.405747901076808, 0)
          );
        }
        // Create the circle
        CreateCircle(
          btr,
          tr,
          new Point3d(endPoint.X - 5.8088, endPoint.Y - 0.3664, 0),
          0.09,
          2,
          false
        );

        // Create the 1
        CreateAndPositionCenteredText(
          tr,
          "1",
          "gmep",
          0.09375,
          1,
          2,
          "0",
          new Point3d(endPoint.X - 5.85897687070053 - 0.145, endPoint.Y - 0.410151417346867, 0)
        );
      }
      else
      {
        var number = 1;
        // for each notes that does not contain *NOT ADDED AS NOTE*, create a note
        foreach (var note in customNotes)
        {
          if (String.IsNullOrEmpty(note))
          {
            continue;
          }
          if (!note.Contains("NOT ADDED AS NOTE"))
          {
            // Create the circle
            CreateCircle(
              btr,
              tr,
              new Point3d(endPoint.X - 5.8088, endPoint.Y - 0.3664 - (lines_of_text * 0.1872), 0),
              0.09,
              2,
              false
            );

            // Create the number
            CreateAndPositionCenteredText(
              tr,
              number.ToString(), // TODO add check for custom note number i.e. one from the database/design tool
              "gmep",
              0.09375,
              1,
              2,
              "0",
              new Point3d(
                endPoint.X - 5.85897687070053 - 0.145,
                endPoint.Y - 0.410151417346867 - (lines_of_text * 0.1872),
                0
              )
            );

            // if the string is longer than 65 characters, find the end of the word closest to the 65th character and split the string there into two strings and check the next string if it is longer than 65 characters and do the same, then return all the strings
            var noteStrings = SplitStringIntoLines(note, 65);
            foreach (var noteString in noteStrings)
            {
              CreateAndPositionText(
                tr,
                noteString,
                "gmep",
                0.09375,
                1,
                2,
                "0",
                new Point3d(
                  endPoint.X - 5.61904201783966,
                  endPoint.Y - 0.405747901076808 - (lines_of_text * 0.1872),
                  0
                )
              );
              lines_of_text++;
            }

            number++;
          }
        }
        if (lines_of_text > lines - 1)
        {
          lines = lines_of_text + 1;
          decrease_y = (lines - 5) * 0.1872;
        }
      }
      for (int i = 0; i <= lines; i++)
      {
        yOffset = y_initial + (i * y_increment);
        CreateLine(
          tr,
          btr,
          endPoint.X,
          endPoint.Y - yOffset,
          endPoint.X - 6.07359999999994,
          endPoint.Y - yOffset,
          "0"
        );
      }

      // Create the vertical lines
      CreateLine(
        tr,
        btr,
        endPoint.X,
        endPoint.Y - 0.0846396524177919,
        endPoint.X,
        endPoint.Y - yOffset,
        "0"
      );
      CreateLine(
        tr,
        btr,
        endPoint.X - 6.07359999999994,
        endPoint.Y - yOffset,
        endPoint.X - 6.07359999999994,
        endPoint.Y + -0.0846396524177919,
        "0"
      );

      return decrease_y;
    }

    private List<string> SplitStringIntoLines(string str, int maxLength)
    {
      List<string> lines = new List<string>();

      while (str.Length > maxLength)
      {
        int index = str.LastIndexOf(' ', maxLength);
        string line = str.Substring(0, index);
        lines.Add(line);
        str = str.Substring(index + 1);
      }

      lines.Add(str);

      return lines;
    }

    private void CreateCenterLines(
      BlockTableRecord btr,
      Transaction tr,
      Point3d startPoint,
      Point3d endPoint,
      bool is2Pole
    )
    {
      // Create horizontal line above
      CreateLine(tr, btr, startPoint.X, endPoint.Y + 0.2533, endPoint.X, endPoint.Y + 0.2533, "0");

      if (is2Pole)
      {
        /*
          // Create the slashed line
          CreateLine(
            tr,
            btr,
            endPoint.X - 6.47536463134611,
            endPoint.Y + 0.0841286798547145,
            endPoint.X - 6.3618865297326,
            endPoint.Y + 0.216793591015815,
            "0"
          );
          CreateLine(
            tr,
            btr,
            endPoint.X - 4.541248855498,
            endPoint.Y + 0.0739331046861764,
            endPoint.X - 4.42777075388449,
            endPoint.Y + 0.206598015847277,
            "0"
          );
          */
        // Create the vertical center lines
        CreateLine(
          tr,
          btr,
          endPoint.X - 4.67999999999984,
          endPoint.Y + 0.277148912917994,
          endPoint.X - 4.67999999999984,
          startPoint.Y - 0.7488,
          "0"
        );
        CreateLine(
          tr,
          btr,
          endPoint.X - 4.30668381424084,
          endPoint.Y + 0.277148912917994,
          endPoint.X - 4.30668381424084,
          startPoint.Y - 0.7488,
          "0"
        );
      }
      else
      {
        /*
        // Create the slashed line
        CreateLine(
          tr,
          btr,
          endPoint.X - 6.47536463134611,
          endPoint.Y + 0.0663322399757078,
          endPoint.X - 6.36188652973283,
          endPoint.Y + 0.198997151136808,
          "0"
        );
        CreateLine(
          tr,
          btr,
          endPoint.X - 4.63023473191822,
          endPoint.Y + 0.0663322399757078,
          endPoint.X - 4.51675663030494,
          endPoint.Y + 0.198997151136808,
          "0"
        );
        CreateLine(
          tr,
          btr,
          endPoint.X - 2.06710862405464,
          endPoint.Y + 0.0663322399757078,
          endPoint.X - 1.95363052244136,
          endPoint.Y + 0.198997151136808,
          "0"
        );
        */
        // Create the vertical center lines
        CreateLine(
          tr,
          btr,
          endPoint.X - 4.30559999999991,
          endPoint.Y + 0.253292187434056,
          endPoint.X - 4.30559999999991,
          startPoint.Y - 0.7488,
          "0"
        );
        CreateLine(
          tr,
          btr,
          endPoint.X - 4.49279999999999,
          endPoint.Y + 0.253292187434056,
          endPoint.X - 4.49279999999999,
          startPoint.Y - 0.7488,
          "0"
        );
        CreateLine(
          tr,
          btr,
          endPoint.X - 4.67999999999984,
          endPoint.Y + 0.253292187434056,
          endPoint.X - 4.67999999999984,
          startPoint.Y - 0.7488,
          "0"
        );
      }

      // Create the circle and lines in the center of the panel
      CreateCenterLinePattern(tr, btr, startPoint, endPoint, is2Pole);
    }

    private void CreateCenterLinePattern(
      Transaction tr,
      BlockTableRecord btr,
      Point3d startPoint,
      Point3d endPoint,
      bool is2Pole
    )
    {
      double maxY = endPoint.Y + 0.2533;
      double increaseY = 0.1872;
      double increaseX = 0.1872;
      double baseX = startPoint.X + 4.3056;
      double currentX = baseX;
      double currentY = startPoint.Y - 0.8424;
      int num = 3;

      if (is2Pole)
      {
        baseX = startPoint.X + 4.30560000000014;
        currentX = baseX;
        increaseX = 0.3733;
        num = 2;
      }

      bool conditionMet = false;

      while (currentY >= maxY && !conditionMet)
      {
        for (int i = 0; i < num; i++)
        {
          if (currentY < maxY)
          {
            conditionMet = true;
            break;
          }

          // Create the center line circles
          CreateCircle(btr, tr, new Point3d(currentX, currentY, 0), 0.0312, 7);

          // Create the horizontal center lines
          CreateLine(
            tr,
            btr,
            startPoint.X + 4.22905693965708,
            currentY,
            startPoint.X + 4.75654306034312,
            currentY,
            "0"
          );

          currentX += increaseX;
          currentY -= increaseY;
        }

        // reset x value
        currentX = baseX;
      }
    }

    public void CreateCircle(
      BlockTableRecord btr,
      Transaction tr,
      Point3d center,
      double radius,
      int colorIndex,
      bool doHatch = true
    )
    {
      using (Circle circle = new Circle())
      {
        circle.Center = center;
        circle.Radius = radius;
        circle.ColorIndex = colorIndex; // Setting the color
        circle.Layer = "0"; // Setting the layer to "0"
        btr.AppendEntity(circle);
        tr.AddNewlyCreatedDBObject(circle, true);

        if (doHatch)
        {
          // Creating a Hatch
          using (Hatch hatch = new Hatch())
          {
            hatch.Layer = "0"; // Setting the layer to "0"
            btr.AppendEntity(hatch);
            tr.AddNewlyCreatedDBObject(hatch, true);
            hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
            hatch.Associative = true;

            // Associating the hatch with the circle
            ObjectIdCollection objIds = new ObjectIdCollection();
            objIds.Add(circle.ObjectId);
            hatch.AppendLoop(HatchLoopTypes.Default, objIds);
            hatch.EvaluateHatch(true);
          }
        }
      }
    }

    private void CreateFooterText(
      Transaction tr,
      Point3d endPoint,
      Dictionary<string, object> panelData,
      bool is2Pole
    )
    {
      if (is2Pole)
      {
        CreateAndPositionText(
          tr,
          "SUB-TOTAL",
          "gmep",
          0.1248,
          0.75,
          256,
          "0",
          new Point3d(endPoint.X - 8.91077927366177, endPoint.Y + 0.0697891578365528, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "A",
          "gmep",
          0.1248,
          0.75,
          256,
          "0",
          new Point3d(endPoint.X - 6.45042438923338, endPoint.Y + 0.0920885745242259, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "B",
          "gmep",
          0.1248,
          0.75,
          256,
          "0",
          new Point3d(endPoint.X - 4.51630861338526, endPoint.Y + 0.0818929993556878, 0)
        );
        CreateAndPositionText(
          tr,
          "=",
          "gmep",
          0.1248,
          0.75,
          256,
          "0",
          new Point3d(endPoint.X - 6.24591440390805, endPoint.Y + 0.0920885745242259, 0)
        );
        CreateAndPositionText(
          tr,
          "=",
          "gmep",
          0.1248,
          0.75,
          256,
          "0",
          new Point3d(endPoint.X - 4.32551576122205, endPoint.Y + 0.0818929993556878, 0)
        );
        CreateAndPositionText(
          tr,
          panelData["subtotal_a"] as string + "VA",
          "gmep",
          0.09375,
          1,
          2,
          "0",
          new Point3d(endPoint.X - 6.08199405082502, endPoint.Y + 0.108280531454625, 0)
        );
        CreateAndPositionText(
          tr,
          panelData["subtotal_b"] as string + "VA",
          "gmep",
          0.09375,
          1,
          2,
          "0",
          new Point3d(endPoint.X - 4.15962890179264, endPoint.Y + 0.0980849562860868, 0)
        );
      }
      else
      {
        CreateAndPositionText(
          tr,
          "SUB-TOTAL",
          "gmep",
          0.1248,
          1,
          7,
          "0",
          new Point3d(endPoint.X - 8.91077927366155, endPoint.Y + 0.0689855381989162, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "A",
          "gmep",
          0.1248,
          0.75,
          7,
          "0",
          new Point3d(endPoint.X - 6.45042438923338, endPoint.Y + 0.0742921346453898, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "B",
          "gmep",
          0.1248,
          0.75,
          7,
          "0",
          new Point3d(endPoint.X - 4.60529448980549, endPoint.Y + 0.0742921346453898, 0)
        );
        CreateAndPositionText(
          tr,
          "\u03A6" + "C",
          "gmep",
          0.1248,
          0.75,
          7,
          "0",
          new Point3d(endPoint.X - 2.04216838194191, endPoint.Y + 0.0742921346453898, 0)
        );
        CreateAndPositionText(
          tr,
          panelData["subtotal_a"] as string + "VA",
          "gmep",
          0.09375,
          1,
          2,
          "0",
          new Point3d(endPoint.X - 6.07732066030258, endPoint.Y + 0.0948263267698053, 0)
        );
        CreateAndPositionText(
          tr,
          panelData["subtotal_b"] as string + "VA",
          "gmep",
          0.09375,
          1,
          2,
          "0",
          new Point3d(endPoint.X - 4.23219076087469, endPoint.Y + 0.0948263267698053, 0)
        );
        CreateAndPositionText(
          tr,
          panelData["subtotal_c"] as string + "VA",
          "gmep",
          0.09375,
          1,
          2,
          "0",
          new Point3d(endPoint.X - 1.66906465301099, endPoint.Y + 0.0948263267698053, 0)
        );
        CreateAndPositionText(
          tr,
          "=",
          "gmep",
          0.1248,
          0.75,
          256,
          "0",
          new Point3d(endPoint.X - 6.24591440390827, endPoint.Y + 0.0742921346453898, 0)
        );
        CreateAndPositionText(
          tr,
          "=",
          "gmep",
          0.1248,
          0.75,
          256,
          "0",
          new Point3d(endPoint.X - 4.40078450448038, endPoint.Y + 0.0742921346453898, 0)
        );
        CreateAndPositionText(
          tr,
          "=",
          "gmep",
          0.1248,
          0.75,
          256,
          "0",
          new Point3d(endPoint.X - 1.8376583966168, endPoint.Y + 0.0742921346453898, 0)
        );
      }
    }

    private double GetEndOfDataY(List<string> list, Point3d startPoint)
    {
      var rowHeight = 0.1872;
      var headerHeight = 0.7488;
      return startPoint.Y - (headerHeight + (rowHeight * ((list.Count + 1) / 2)));
    }

    private void CreateRectangle(
      BlockTableRecord btr,
      Transaction tr,
      Point3d topRightCorner,
      Point3d startPoint,
      Point3d endPoint,
      string layerName
    )
    {
      // Create the rectangle
      var rect = new Autodesk.AutoCAD.DatabaseServices.Polyline(4);
      rect.AddVertexAt(0, new Point2d(startPoint.X, startPoint.Y), 0, 0, 0);
      rect.AddVertexAt(1, new Point2d(startPoint.X, endPoint.Y), 0, 0, 0);
      rect.AddVertexAt(2, new Point2d(endPoint.X, endPoint.Y), 0, 0, 0);
      rect.AddVertexAt(3, new Point2d(endPoint.X, startPoint.Y), 0, 0, 0);
      rect.Closed = true;

      // Set the global width property
      rect.ConstantWidth = 0.02;

      // Set the layer to "0"
      rect.Layer = layerName;

      btr.AppendEntity(rect);
      tr.AddNewlyCreatedDBObject(rect, true);
    }

    private static ObjectId CreateText(
      string content,
      string style,
      TextHorizontalMode horizontalMode,
      TextVerticalMode verticalMode,
      double height,
      double widthFactor,
      Autodesk.AutoCAD.Colors.Color color,
      string layer
    )
    {
      var (doc, db, _) = PanelCommands.GetGlobals();

      // Check if the layer exists
      using (var tr = db.TransactionManager.StartTransaction())
      {
        var layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

        if (!layerTable.Has(layer))
        {
          // Layer doesn't exist, create it
          var newLayer = new LayerTableRecord();
          newLayer.Name = layer;

          layerTable.UpgradeOpen();
          layerTable.Add(newLayer);
          tr.AddNewlyCreatedDBObject(newLayer, true);
        }

        tr.Commit();
      }

      using (var tr = doc.TransactionManager.StartTransaction())
      {
        var textStyleId = GetTextStyleId(style);
        var textStyle = (TextStyleTableRecord)tr.GetObject(textStyleId, OpenMode.ForRead);

        if (textStyle.FileName.ToLower().Contains("architxt"))
        {
          if (widthFactor > 0.85)
          {
            widthFactor = 0.85;
          }
          content = content.Replace("\u03A6", "\u0081");
        }

        var text = new DBText
        {
          TextString = content,
          Height = height,
          WidthFactor = widthFactor,
          Color = color,
          Layer = layer,
          TextStyleId = textStyleId,
          HorizontalMode = horizontalMode,
          VerticalMode = verticalMode,
          Justify = AttachmentPoint.BaseLeft,
        };

        var currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
        currentSpace.AppendEntity(text);
        tr.AddNewlyCreatedDBObject(text, true);

        tr.Commit();

        return text.ObjectId;
      }
    }

    private static void CreateAndPositionText(
      Transaction tr,
      string content,
      string style,
      double height,
      double widthFactor,
      int colorIndex,
      string layerName,
      Point3d position,
      TextHorizontalMode horizontalMode = TextHorizontalMode.TextLeft,
      TextVerticalMode verticalMode = TextVerticalMode.TextBase
    )
    {
      var color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
        Autodesk.AutoCAD.Colors.ColorMethod.ByLayer,
        (short)colorIndex
      );
      var textId = CreateText(
        content,
        style,
        horizontalMode,
        verticalMode,
        height,
        widthFactor,
        color,
        layerName
      );
      var text = (DBText)tr.GetObject(textId, OpenMode.ForWrite);
      text.Position = position;
    }

    private void CreateAndPositionFittedText(
      Transaction tr,
      string content,
      string style,
      double height,
      double widthFactor,
      int colorIndex,
      string layerName,
      Point3d position,
      double length,
      TextHorizontalMode horizontalMode = TextHorizontalMode.TextLeft,
      TextVerticalMode verticalMode = TextVerticalMode.TextBase
    )
    {
      var color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
        Autodesk.AutoCAD.Colors.ColorMethod.ByLayer,
        (short)colorIndex
      );
      var textId = CreateText(
        content,
        style,
        horizontalMode,
        verticalMode,
        height,
        widthFactor,
        color,
        layerName
      );
      var text = (DBText)tr.GetObject(textId, OpenMode.ForWrite);

      double naturalWidth = text.GeometricExtents.MaxPoint.X - text.GeometricExtents.MinPoint.X;
      text.WidthFactor = length / naturalWidth; // This will stretch or squeeze text to fit between points
      text.Position = position;
    }

    private void CreateAndPositionCenteredText(
      Transaction tr,
      string content,
      string style,
      double height,
      double widthFactor,
      int colorIndex,
      string layerName,
      Point3d position,
      TextHorizontalMode horizontalMode = TextHorizontalMode.TextLeft,
      TextVerticalMode verticalMode = TextVerticalMode.TextBase
    )
    {
      var color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
        Autodesk.AutoCAD.Colors.ColorMethod.ByLayer,
        (short)colorIndex
      );
      var textId = CreateText(
        content,
        style,
        horizontalMode,
        verticalMode,
        height,
        widthFactor,
        color,
        layerName
      );
      var text = (DBText)tr.GetObject(textId, OpenMode.ForWrite);
      text.Justify = AttachmentPoint.BaseCenter;
      double x = position.X;
      text.AlignmentPoint = new Point3d(x + 0.1903, position.Y, 0);
    }

    private void CreateAndPositionRightText(
      Transaction tr,
      string content,
      string style,
      double height,
      double widthFactor,
      int colorIndex,
      string layerName,
      Point3d position,
      TextHorizontalMode horizontalMode = TextHorizontalMode.TextLeft,
      TextVerticalMode verticalMode = TextVerticalMode.TextBase
    )
    {
      var color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(
        Autodesk.AutoCAD.Colors.ColorMethod.ByLayer,
        (short)colorIndex
      );
      var textId = CreateText(
        content,
        style,
        horizontalMode,
        verticalMode,
        height,
        widthFactor,
        color,
        layerName
      );
      var text = (DBText)tr.GetObject(textId, OpenMode.ForWrite);
      text.Justify = AttachmentPoint.BaseRight;
      double x = position.X;
      text.AlignmentPoint = new Point3d(x + 0.46, position.Y, 0);
    }

    private void CreateVerticalLines(
      Transaction tr,
      BlockTableRecord btr,
      Point3d startPoint,
      double[] distances,
      double startY,
      double endY,
      string layer
    )
    {
      foreach (double distance in distances)
      {
        var lineStart = new Point3d(startPoint.X + distance, startY, 0);
        var lineEnd = new Point3d(startPoint.X + distance, endY, 0);
        var line = new Line(lineStart, lineEnd);
        line.Layer = layer;

        btr.AppendEntity(line);
        tr.AddNewlyCreatedDBObject(line, true);
      }
    }

    private void CreateLines(
      Transaction tr,
      BlockTableRecord btr,
      IEnumerable<(double startX, double startY, double endX, double endY, string layer)> lines
    )
    {
      foreach (var (startX, startY, endX, endY, layer) in lines)
      {
        CreateLine(tr, btr, startX, startY, endX, endY, layer);
      }
    }

    public static void CreateLine(
      Transaction tr,
      BlockTableRecord btr,
      double startX,
      double startY,
      double endX,
      double endY,
      string layer
    )
    {
      var lineStart = new Point3d(startX, startY, 0);
      var lineEnd = new Point3d(endX, endY, 0);
      var line = new Line(lineStart, lineEnd);
      line.Layer = layer;

      btr.AppendEntity(line);
      tr.AddNewlyCreatedDBObject(line, true);
    }

    private int ProcessTextData(
      Transaction tr,
      BlockTableRecord btr,
      Point3d startPoint,
      Dictionary<string, object> panelData,
      bool is2Pole
    )
    {
      List<bool> leftBreakersHighlight = (List<bool>)panelData["breaker_left_highlights"];
      List<bool> rightBreakersHighlight = (List<bool>)panelData["breaker_right_highlights"];

      var largest_level = 0;

      if (!is2Pole)
      {
        ProcessSideData(tr, btr, startPoint, panelData, true, false);
        ProcessSideData(tr, btr, startPoint, panelData, false, false);
      }
      else
      {
        ProcessSideData2P(tr, btr, startPoint, panelData, true, true);
        ProcessSideData2P(tr, btr, startPoint, panelData, false, true);
      }

      InsertKeepBreakers(startPoint, leftBreakersHighlight, true);
      InsertKeepBreakers(startPoint, rightBreakersHighlight, false);

      // if the key "description_left_tags" exists in panelStorage, return null
      if (panelData.ContainsKey("description_left_tags"))
      {
        List<string> descriptionLeftTags = (List<string>)panelData["description_left_tags"];
        List<string> descriptionRightTags = (List<string>)panelData["description_right_tags"];

        List<string> notes = (List<string>)panelData["notes"];

        notes = RemoveNotAddedAsNotes(notes);

        Dictionary<string, List<bool>> leftSide = ConvertTagsAndNotesToDictionary(
          descriptionLeftTags,
          notes,
          (List<string>)panelData["description_left"]
        );
        Dictionary<string, List<bool>> rightSide = ConvertTagsAndNotesToDictionary(
          descriptionRightTags,
          notes,
          (List<string>)panelData["description_right"]
        );

        var left_largest_level = InsertBreakerNotes(startPoint, leftSide, true);
        var right_largest_level = InsertBreakerNotes(startPoint, rightSide, false);

        largest_level =
          left_largest_level > right_largest_level ? left_largest_level : right_largest_level;
      }
      return largest_level;
    }

    private Dictionary<string, List<bool>> ConvertTagsAndNotesToDictionary(
      List<string> tags,
      List<string> notes,
      List<string> descriptions
    )
    {
      Dictionary<string, List<bool>> notesWithBools = new Dictionary<string, List<bool>>();

      foreach (string note in notes)
      {
        List<bool> bools = new List<bool>();
        foreach (string tag in tags)
        {
          var i = tags.IndexOf(tag);
          if (
            descriptions[i * 2] == "SPACE"
            && note
              == "DENOTES EXISTING CIRCUIT BREAKER TO REMAIN; ALL OTHERS ARE NEW TO MATCH EXISTING."
          )
          {
            bools.Add(false);
          }
          else
          {
            bools.Add(tag.Split('|').Contains(note));
          }
        }
        notesWithBools.Add((notes.IndexOf(note) + 1).ToString(), bools);
      }
      return notesWithBools;
    }

    private List<string> RemoveNotAddedAsNotes(List<string> notes)
    {
      List<string> newNotes = new List<string>();
      foreach (string note in notes)
      {
        if (String.IsNullOrEmpty(note))
          continue;
        if (!note.Contains("NOT ADDED AS NOTE"))
        {
          newNotes.Add(note);
        }
      }
      return newNotes;
    }

    private Dictionary<string, List<int>> ConvertBooleansToLevels(
      Dictionary<string, List<bool>> notesWithBools
    )
    {
      Dictionary<string, List<int>> notesWithLevels = new Dictionary<string, List<int>>();
      List<int> maxLevelsSoFar = Enumerable.Repeat(0, notesWithBools.First().Value.Count).ToList();

      foreach (var pair in notesWithBools)
      {
        List<int> currentLevels = new List<int>();

        for (int i = 0; i < pair.Value.Count; i++)
        {
          if (pair.Value[i])
          {
            int level = maxLevelsSoFar[i] + 1;
            currentLevels.Add(level);
            maxLevelsSoFar[i] = level;
          }
          else
          {
            currentLevels.Add(0);
          }
        }

        notesWithLevels.Add(pair.Key, currentLevels);
      }

      return notesWithLevels;
    }

    private int InsertBreakerNotes(
      Point3d startPoint,
      Dictionary<string, List<bool>> notesWithBools,
      bool left
    )
    {
      if (notesWithBools.Count == 0)
      {
        return 0;
      }

      Point3d botPoint = new Point3d(0, 0, 0);
      Point3d topPoint = new Point3d(0, 0, 0);

      double header_height = 0.7488;
      double panel_width = 8.9856;
      double row_height = 0.1872;
      double start_x = startPoint.X + (left ? 0 : panel_width);
      double start_y = startPoint.Y - header_height;
      double displacement = left ? -1 : 1;

      bool currentlyKeeping = false;

      var notesWithLevels = ConvertBooleansToLevels(notesWithBools);

      foreach (var item in notesWithLevels)
      {
        int currentLevel = 0;
        int previousLevel = 0;
        for (int i = 0; i < item.Value.Count; i += 1)
        {
          currentLevel = item.Value[i];
          if (currentLevel > 0)
          {
            if (currentLevel != previousLevel && previousLevel > 0)
            {
              botPoint = new Point3d(
                start_x + ((item.Value[i] - 1) * (left ? -1 : 1) * 0.2),
                start_y - (row_height * (i)),
                0
              );
              KeepBreakersGivenPoints(
                topPoint,
                botPoint,
                new Point3d(topPoint.X + displacement, topPoint.Y, 0),
                item.Key
              );
              topPoint = new Point3d(
                start_x + ((item.Value[i] - 1) * (left ? -1 : 1) * 0.2),
                start_y - (row_height * (i)),
                0
              );
            }
            if (!currentlyKeeping)
            {
              topPoint = new Point3d(
                start_x + ((item.Value[i] - 1) * (left ? -1 : 1) * 0.2),
                start_y - (row_height * (i)),
                0
              );
              currentlyKeeping = true;
            }
            if (i >= item.Value.Count - 1)
            {
              botPoint = new Point3d(
                start_x + ((item.Value[i] - 1) * (left ? -1 : 1) * 0.2),
                start_y - (row_height * (i + 1)),
                0
              );
              KeepBreakersGivenPoints(
                topPoint,
                botPoint,
                new Point3d(topPoint.X + displacement, topPoint.Y, 0),
                item.Key
              );
            }
          }
          else if (currentlyKeeping)
          {
            botPoint = new Point3d(
              start_x + ((item.Value[i] - 1) * (left ? -1 : 1) * 0.2),
              start_y - (row_height * i),
              0
            );
            currentlyKeeping = false;
            KeepBreakersGivenPoints(
              topPoint,
              botPoint,
              new Point3d(topPoint.X + displacement, topPoint.Y, 0),
              item.Key
            );
          }
          previousLevel = currentLevel;
        }
        currentlyKeeping = false;
      }

      return GetLargestLevel(notesWithLevels);
    }

    private int GetLargestLevel(Dictionary<string, List<int>> notesWithLevels)
    {
      int largestLevel = 0;
      foreach (var item in notesWithLevels)
      {
        foreach (int level in item.Value)
        {
          if (level > largestLevel)
          {
            largestLevel = level;
          }
        }
      }
      return largestLevel;
    }

    private void InsertKeepBreakers(Point3d startPoint, List<bool> breakersHighlight, bool left)
    {
      Point3d botPoint = new Point3d(0, 0, 0);
      Point3d topPoint = new Point3d(0, 0, 0);

      double header_height = 0.7488;
      double panel_width = 8.9856;
      double row_height = 0.1872;
      double start_x = startPoint.X + (left ? 0 : panel_width);
      double start_y = startPoint.Y - header_height;
      double displacement = left ? -1 : 1;

      bool currentlyKeeping = false;

      for (int i = 0; i < breakersHighlight.Count; i += 2)
      {
        if (breakersHighlight[i])
        {
          if (!currentlyKeeping)
          {
            topPoint = new Point3d(start_x, start_y - (row_height * (i / 2)), 0);
            currentlyKeeping = true;
          }
          if (i >= breakersHighlight.Count - 2)
          {
            botPoint = new Point3d(start_x, start_y - (row_height * ((i + 2) / 2)), 0);
            KeepBreakersGivenPoints(
              topPoint,
              botPoint,
              new Point3d(topPoint.X + displacement, topPoint.Y, 0),
              "1"
            );
          }
        }
        else if (currentlyKeeping)
        {
          botPoint = new Point3d(start_x, start_y - (row_height * (i / 2)), 0);
          currentlyKeeping = false;
          KeepBreakersGivenPoints(
            topPoint,
            botPoint,
            new Point3d(topPoint.X + displacement, topPoint.Y, 0),
            "1"
          );
        }
      }
    }

    private (
      List<string>,
      List<string>,
      List<string>,
      List<string>,
      List<string>,
      List<string>,
      List<bool>,
      List<string>
    ) GetCorrectBreakerData(Dictionary<string, object> panelData, bool left, bool is2Pole)
    {
      var descriptions = new List<string>();
      var breakers = new List<string>();
      var circuits = new List<string>();
      var phaseA = new List<string>();
      var phaseB = new List<string>();
      var phaseC = new List<string>();
      var descriptionHighlights = new List<bool>();
      var descriptionTags = new List<string>();

      if (left)
      {
        descriptions = panelData["description_left"] as List<string>;
        breakers = panelData["breaker_left"] as List<string>;
        circuits = panelData["circuit_left"] as List<string>;
        phaseA = panelData["phase_a_left"] as List<string>;
        phaseB = panelData["phase_b_left"] as List<string>;
        if (!is2Pole)
        {
          phaseC = panelData["phase_c_left"] as List<string>;
        }
        descriptionHighlights = panelData["description_left_highlights"] as List<bool>;
        if (panelData.ContainsKey("description_left_tags"))
        {
          descriptionTags = panelData["description_left_tags"] as List<string>;
        }
      }
      else
      {
        descriptions = panelData["description_right"] as List<string>;
        breakers = panelData["breaker_right"] as List<string>;
        circuits = panelData["circuit_right"] as List<string>;
        phaseA = panelData["phase_a_right"] as List<string>;
        phaseB = panelData["phase_b_right"] as List<string>;
        if (!is2Pole)
        {
          phaseC = panelData["phase_c_right"] as List<string>;
        }
        descriptionHighlights = panelData["description_right_highlights"] as List<bool>;
        if (panelData.ContainsKey("description_right_tags"))
        {
          descriptionTags = panelData["description_right_tags"] as List<string>;
        }
      }

      return (
        descriptions,
        breakers,
        circuits,
        phaseA,
        phaseB,
        phaseC,
        descriptionHighlights,
        descriptionTags
      );
    }

    private void ProcessSideData(
      Transaction tr,
      BlockTableRecord btr,
      Point3d startPoint,
      Dictionary<string, object> panelData,
      bool left,
      bool is2Pole
    )
    {
      var (descriptions, breakers, circuits, _, _, _, _, _) = GetCorrectBreakerData(
        panelData,
        left,
        is2Pole
      );

      Dictionary<string, double> data = new Dictionary<string, double>();

      data.Add("row height y", 0.1872);
      data.Add("half row height y", 0.0936);
      data.Add("initial half breaker text y", -0.816333638994546);
      data.Add("header height", 0.7488);

      int i = 0;
      while (i < descriptions.Count)
      {
        BreakerType breakerType = GetBreakerType(breakers, circuits, i);
        if (breakerType == BreakerType.SplitMini)
        {
          CreateHalfSplit2PoleBreakers(tr, btr, startPoint, panelData, data, left, is2Pole, i);
          i += 4;
        }
        else if (breakerType == BreakerType.Mini)
        {
          CreateHalfBreaker(tr, btr, startPoint, panelData, data, left, is2Pole, i);
          i += 2;
        }
        else if (breakerType == BreakerType.SinglePole)
        {
          Create1PoleBreaker(tr, btr, startPoint, panelData, data, left, is2Pole, i);
          i += 2;
        }
        else if (breakerType == BreakerType.TwoPole)
        {
          Create2PoleBreaker(tr, btr, startPoint, panelData, data, left, is2Pole, i);
          i += 4;
        }
        else
        {
          Create3PoleBreaker(tr, btr, startPoint, panelData, data, left, is2Pole, i);
          i += 6;
        }
      }
    }

    private void ProcessSideData2P(
      Transaction tr,
      BlockTableRecord btr,
      Point3d startPoint,
      Dictionary<string, object> panelData,
      bool left,
      bool is2Pole
    )
    {
      var (descriptions, breakers, circuits, _, _, _, _, _) = GetCorrectBreakerData(
        panelData,
        left,
        is2Pole
      );

      Dictionary<string, double> data = new Dictionary<string, double>
      {
        { "row height y", 0.1872 },
        { "half row height y", 0.0936 },
        { "initial half breaker text y", -0.816333638994546 },
        { "header height", 0.7488 },
      };
      int i = 0;
      while (i < descriptions.Count)
      {
        BreakerType breakerType = GetBreakerType(breakers, circuits, i);
        if (breakerType == BreakerType.SplitMini)
        {
          CreateHalfSplit2PoleBreakers2Ph(tr, btr, startPoint, panelData, data, left, is2Pole, i);
          i += 4;
        }
        else if (breakerType == BreakerType.Mini)
        {
          CreateHalfBreaker2P(tr, btr, startPoint, panelData, data, left, is2Pole, i);
          i += 2;
        }
        else if (breakerType == BreakerType.SinglePole)
        {
          Create1PoleBreaker2P(tr, btr, startPoint, panelData, data, left, is2Pole, i);
          i += 2;
        }
        else
        {
          Create2PoleBreaker2P(tr, btr, startPoint, panelData, data, left, is2Pole, i);
          i += 4;
        }
      }
    }

    private void CreateHalfBreaker(
      Transaction tr,
      BlockTableRecord btr,
      Point3d startPoint,
      Dictionary<string, object> panelData,
      Dictionary<string, double> data,
      bool left,
      bool is2Pole,
      int i
    )
    {
      var (
        descriptions,
        breakers,
        circuits,
        phaseA,
        phaseB,
        phaseC,
        descriptionHighlights,
        descriptionTags
      ) = GetCorrectBreakerData(panelData, left, is2Pole);

      List<string> phaseList = GetPhaseList(i, phaseA, phaseB, phaseC);

      double descriptionX = GetDescriptionX(startPoint, left);
      double phaseX = GetPhaseX(i, startPoint, left);
      double breakerX = GetBreakerX(startPoint, left);
      double circuitX = GetCircuitX(startPoint, left);
      double textHeight = 0.0725;

      string circuit = circuits[i];

      for (var j = i; j <= i + 1; j++)
      {
        if (Regex.IsMatch(descriptions[j], @"^[a-fA-F0-9]{8}(-[a-fA-F0-9]{4}){3}-[a-fA-F0-9]{12}$"))
        {
          List<Dictionary<string, object>> panelStorage = myForm.RetrieveSavedPanelData();
          foreach (Dictionary<string, object> panel in panelStorage)
          {
            if ((panel["id"] as string).ToLower() == descriptions[j].ToLower())
            {
              descriptions[j] = "PANEL " + panel["panel"] as string;
            }
          }
        }
        string description =
          (descriptionHighlights[j] && descriptions[j] != "EXISTING LOAD")
            ? "(E)" + descriptions[j]
            : descriptions[j];
        string breaker = breakers[j] + "-1";
        string phase = phaseList[j];
        circuit = circuits[j];
        double height = startPoint.Y + (-0.831333638994546 - ((double)j / 2 * 0.1872));

        CreateAndPositionText(
          tr,
          description.Trim(),
          "gmep",
          textHeight,
          1.0,
          2,
          "0",
          new Point3d(descriptionX, height, 0)
        );
        if (phase != "0")
          CreateAndPositionCenteredText(
            tr,
            phase,
            "gmep",
            textHeight,
            1.0,
            2,
            "0",
            new Point3d(phaseX, height, 0)
          );
        CreateAndPositionText(
          tr,
          breaker,
          "gmep",
          textHeight,
          1.0,
          2,
          "0",
          new Point3d(breakerX, height, 0)
        );
        CreateAndPositionText(
          tr,
          circuit,
          "gmep",
          textHeight,
          1.0,
          7,
          "0",
          new Point3d(circuitX, height, 0)
        );
      }

      CreateHorizontalLine(startPoint.X, startPoint.Y, circuit, left, tr, btr);
    }

    private void CreateHalfBreaker2P(
      Transaction tr,
      BlockTableRecord btr,
      Point3d startPoint,
      Dictionary<string, object> panelData,
      Dictionary<string, double> data,
      bool left,
      bool is2Pole,
      int i
    )
    {
      var (
        descriptions,
        breakers,
        circuits,
        phaseA,
        phaseB,
        phaseC,
        descriptionHighlights,
        descriptionTags
      ) = GetCorrectBreakerData(panelData, left, is2Pole);

      List<string> phaseList = GetPhaseList2P(i, phaseA, phaseB);

      double descriptionX = GetDescriptionX2P(startPoint, left);
      double phaseX = GetPhaseX2P(i, startPoint, left);
      double breakerX = GetBreakerX(startPoint, left);
      double circuitX = GetCircuitX(startPoint, left);
      double textHeight = 0.0725;

      string circuit = circuits[i];

      for (var j = i; j <= i + 1; j++)
      {
        if (Regex.IsMatch(descriptions[j], @"^[a-fA-F0-9]{8}(-[a-fA-F0-9]{4}){3}-[a-fA-F0-9]{12}$"))
        {
          List<Dictionary<string, object>> panelStorage = myForm.RetrieveSavedPanelData();
          foreach (Dictionary<string, object> panel in panelStorage)
          {
            if ((panel["id"] as string).ToLower() == descriptions[j].ToLower())
            {
              descriptions[j] = "PANEL " + panel["panel"] as string;
            }
          }
        }
        string description =
          (descriptionHighlights[j] && descriptions[j] != "EXISTING LOAD")
            ? "(E)" + descriptions[j]
            : descriptions[j];
        string breaker = breakers[j] + "-1";
        string phase = phaseList[j];
        circuit = circuits[j];
        double height = startPoint.Y + (-0.831333638994546 - ((double)j / 2 * 0.1872));

        CreateAndPositionText(
          tr,
          description.Trim(),
          "gmep",
          textHeight,
          1.0,
          2,
          "0",
          new Point3d(descriptionX, height, 0)
        );
        if (phase != "0")
          CreateAndPositionCenteredText(
            tr,
            phase,
            "gmep",
            textHeight,
            1.0,
            2,
            "0",
            new Point3d(phaseX, height, 0)
          );
        CreateAndPositionText(
          tr,
          breaker,
          "gmep",
          textHeight,
          1.0,
          2,
          "0",
          new Point3d(breakerX, height, 0)
        );
        CreateAndPositionText(
          tr,
          circuit,
          "gmep",
          textHeight,
          1.0,
          7,
          "0",
          new Point3d(circuitX, height, 0)
        );
      }

      CreateHorizontalLine(startPoint.X, startPoint.Y, circuit, left, tr, btr);
    }

    private void Create1PoleBreaker(
      Transaction tr,
      BlockTableRecord btr,
      Point3d startPoint,
      Dictionary<string, object> panelData,
      Dictionary<string, double> data,
      bool left,
      bool is2Pole,
      int i
    )
    {
      var (descriptions, breakers, circuits, phaseA, phaseB, phaseC, descriptionHighlights, _) =
        GetCorrectBreakerData(panelData, left, is2Pole);

      List<string> phaseList = GetPhaseList(i, phaseA, phaseB, phaseC);
      if (Regex.IsMatch(descriptions[i], @"^[a-fA-F0-9]{8}(-[a-fA-F0-9]{4}){3}-[a-fA-F0-9]{12}$"))
      {
        List<Dictionary<string, object>> panelStorage = myForm.RetrieveSavedPanelData();
        foreach (Dictionary<string, object> panel in panelStorage)
        {
          if ((panel["id"] as string).ToLower() == descriptions[i].ToLower())
          {
            descriptions[i] = "PANEL " + panel["panel"] as string;
          }
        }
      }
      string description =
        (descriptionHighlights[i] && descriptions[i] != "EXISTING LOAD")
          ? "(E)" + descriptions[i]
          : descriptions[i];
      string breaker = breakers[i] + "-1";
      string phase = phaseList[i];
      string circuit = circuits[i];
      double height = startPoint.Y + (-0.890211813771344 - ((i / 2) * 0.1872));
      double descriptionX = GetDescriptionX(startPoint, left);
      double phaseX = GetPhaseX(i, startPoint, left);
      double breakerX = GetBreakerX(startPoint, left);
      double circuitX = GetCircuitX(startPoint, left);
      double length = 0.2300;

      CreateAndPositionText(
        tr,
        description.Trim(),
        "gmep",
        0.09375,
        GetDescriptionWidthFactor(description),
        2,
        "0",
        new Point3d(descriptionX, height, 0)
      );
      if (phase != "0")
        CreateAndPositionCenteredText(
          tr,
          phase,
          "gmep",
          0.09375,
          1.0,
          2,
          "0",
          new Point3d(phaseX, height, 0)
        );
      if (breaker != "-1")
        CreateAndPositionFittedText(
          tr,
          breaker,
          "gmep",
          0.09375,
          1.0,
          2,
          "0",
          new Point3d(breakerX, height, 0),
          length
        );
      CreateAndPositionText(
        tr,
        circuit,
        "gmep",
        0.09375,
        1.0,
        7,
        "0",
        new Point3d(circuitX, height, 0)
      );
      CreateHorizontalLine(startPoint.X, startPoint.Y, circuit, left, tr, btr);
    }

    private void Create1PoleBreaker2P(
      Transaction tr,
      BlockTableRecord btr,
      Point3d startPoint,
      Dictionary<string, object> panelData,
      Dictionary<string, double> data,
      bool left,
      bool is2Pole,
      int i
    )
    {
      var (
        descriptions,
        breakers,
        circuits,
        phaseA,
        phaseB,
        phaseC,
        descriptionHighlights,
        descriptionTags
      ) = GetCorrectBreakerData(panelData, left, is2Pole);

      List<string> phaseList = GetPhaseList2P(i, phaseA, phaseB);
      if (Regex.IsMatch(descriptions[i], @"^[a-fA-F0-9]{8}(-[a-fA-F0-9]{4}){3}-[a-fA-F0-9]{12}$"))
      {
        List<Dictionary<string, object>> panelStorage = myForm.RetrieveSavedPanelData();
        foreach (Dictionary<string, object> panel in panelStorage)
        {
          if ((panel["id"] as string).ToLower() == descriptions[i].ToLower())
          {
            descriptions[i] = "PANEL " + panel["panel"] as string;
          }
        }
      }
      string description =
        (descriptionHighlights[i] && descriptions[i] != "EXISTING LOAD")
          ? "(E)" + descriptions[i]
          : descriptions[i];
      string breaker = breakers[i] + "-1";
      string phase = phaseList[i];
      string circuit = circuits[i];
      double height = startPoint.Y + (-0.890211813771344 - ((i / 2) * 0.1872));
      double descriptionX = GetDescriptionX2P(startPoint, left);
      double phaseX = GetPhaseX2P(i, startPoint, left);
      double breakerX = GetBreakerX(startPoint, left);
      double circuitX = GetCircuitX(startPoint, left);
      double length = 0.2300;

      CreateAndPositionText(
        tr,
        description.Trim(),
        "gmep",
        0.09375,
        GetDescriptionWidthFactor(description),
        2,
        "0",
        new Point3d(descriptionX, height, 0)
      );
      if (phase != "0")
        CreateAndPositionCenteredText(
          tr,
          phase,
          "gmep",
          0.09375,
          1.0,
          2,
          "0",
          new Point3d(phaseX, height, 0)
        );
      if (breaker != "-1")
        CreateAndPositionFittedText(
          tr,
          breaker,
          "gmep",
          0.09375,
          1.0,
          2,
          "0",
          new Point3d(breakerX, height, 0),
          length
        );
      CreateAndPositionText(
        tr,
        circuit,
        "gmep",
        0.09375,
        1.0,
        7,
        "0",
        new Point3d(circuitX, height, 0)
      );
      CreateHorizontalLine(startPoint.X, startPoint.Y, circuit, left, tr, btr);
    }

    private void Create2PoleBreaker(
      Transaction tr,
      BlockTableRecord btr,
      Point3d startPoint,
      Dictionary<string, object> panelData,
      Dictionary<string, double> data,
      bool left,
      bool is2Pole,
      int i
    )
    {
      var (
        descriptions,
        breakers,
        circuits,
        phaseA,
        phaseB,
        phaseC,
        descriptionHighlights,
        descriptionTags
      ) = GetCorrectBreakerData(panelData, left, is2Pole);

      double descriptionX = GetDescriptionX(startPoint, left);
      double breakerX = GetBreakerX(startPoint, left);
      double circuitX = GetCircuitX(startPoint, left);
      double length = 0.14;

      for (var j = i; j <= i + 2; j += 2)
      {
        List<string> phaseList = GetPhaseList(j, phaseA, phaseB, phaseC);
        if (Regex.IsMatch(descriptions[j], @"^[a-fA-F0-9]{8}(-[a-fA-F0-9]{4}){3}-[a-fA-F0-9]{12}$"))
        {
          List<Dictionary<string, object>> panelStorage = myForm.RetrieveSavedPanelData();
          foreach (Dictionary<string, object> panel in panelStorage)
          {
            if ((panel["id"] as string).ToLower() == descriptions[j].ToLower())
            {
              descriptions[j] = "PANEL " + panel["panel"] as string;
            }
          }
        }
        string description =
          (descriptionHighlights[j] && descriptions[j] != "EXISTING LOAD")
            ? "(E)" + descriptions[j]
            : descriptions[j];
        string breaker = breakers[j];
        string phase = phaseList[j];
        string circuit = circuits[j];
        double height = startPoint.Y + (-0.890211813771344 - ((j / 2) * 0.1872));
        double phaseX = GetPhaseX(j, startPoint, left);

        if (j == i + 2)
        {
          description = "---";
          breakerX += 0.16;
          length = 0.07;
        }

        CreateAndPositionText(
          tr,
          description.Trim(),
          "gmep",
          0.09375,
          GetDescriptionWidthFactor(description),
          2,
          "0",
          new Point3d(descriptionX, height, 0)
        );
        if (phase != "0")
          CreateAndPositionCenteredText(
            tr,
            phase,
            "gmep",
            0.09375,
            1.0,
            2,
            "0",
            new Point3d(phaseX, height, 0)
          );
        if (breaker != "")
          CreateAndPositionFittedText(
            tr,
            breaker,
            "gmep",
            0.09375,
            1.0,
            2,
            "0",
            new Point3d(breakerX, height, 0),
            length
          );
        CreateAndPositionText(
          tr,
          circuit,
          "gmep",
          0.09375,
          1.0,
          7,
          "0",
          new Point3d(circuitX, height, 0)
        );
        CreateHorizontalLine(startPoint.X, startPoint.Y, circuit, left, tr, btr);
      }

      CreateBreakerLine(startPoint, i, left, tr, btr, 4);
    }

    private void CreateHalfSplit2PoleBreakers(
      Transaction tr,
      BlockTableRecord btr,
      Point3d startPoint,
      Dictionary<string, object> panelData,
      Dictionary<string, double> data,
      bool left,
      bool is2Pole,
      int i
    )
    {
      var (
        descriptions,
        breakers,
        circuits,
        phaseA,
        phaseB,
        phaseC,
        descriptionHighlights,
        descriptionTags
      ) = GetCorrectBreakerData(panelData, left, is2Pole);

      double descriptionX = GetDescriptionX(startPoint, left);
      double breakerX = GetBreakerX(startPoint, left);
      double circuitX = GetCircuitX(startPoint, left);
      double textHeight = 0.0725;
      double length = 0.14;

      for (var j = i; j <= i + 3; j++)
      {
        List<string> phaseList = GetPhaseList(j, phaseA, phaseB, phaseC);
        if (Regex.IsMatch(descriptions[j], @"^[a-fA-F0-9]{8}(-[a-fA-F0-9]{4}){3}-[a-fA-F0-9]{12}$"))
        {
          List<Dictionary<string, object>> panelStorage = myForm.RetrieveSavedPanelData();
          foreach (Dictionary<string, object> panel in panelStorage)
          {
            if ((panel["id"] as string).ToLower() == descriptions[j].ToLower())
            {
              descriptions[j] = "PANEL " + panel["panel"] as string;
            }
          }
        }
        string description =
          (descriptionHighlights[j] && descriptions[j] != "EXISTING LOAD")
            ? "(E)" + descriptions[j]
            : descriptions[j];
        string breaker = breakers[j];
        string phase = phaseList[j];
        string circuit = circuits[j];
        double height = startPoint.Y + (-0.890211813771344 - ((j / 2) * 0.1872));
        double phaseX = GetPhaseX(j, startPoint, left);

        if (j == i + 2)
        {
          description = "---";
          breakerX += 0.16;
        }
        if (j == i + 3)
        {
          breakerX -= 0.16 + 0.045;
        }

        double yShift = 0;
        if ((j - i) % 2 == 0)
        {
          yShift = 0.0575;
        }
        else
        {
          yShift = -0.0378;
        }
        height = height + yShift;

        if ((j - i) == 0 || (j - i) == 3)
        {
          breaker = breaker + "-1";
          length = 0.23;
        }
        else if ((j - i) == 1)
        {
          length = 0.105;
        }
        else
        {
          length = 0.045;
          breakerX += 0.045;
        }

        CreateAndPositionText(
          tr,
          description.Trim(),
          "gmep",
          textHeight,
          1,
          2,
          "0",
          new Point3d(descriptionX, height, 0)
        );
        if (phase != "0")
          CreateAndPositionCenteredText(
            tr,
            phase,
            "gmep",
            textHeight,
            1.0,
            2,
            "0",
            new Point3d(phaseX, height, 0)
          );
        if (breaker != "")
        {
          CreateAndPositionFittedText(
            tr,
            breaker,
            "gmep",
            textHeight,
            1.0,
            2,
            "0",
            new Point3d(breakerX, height, 0),
            length
          );
        }
        CreateAndPositionText(
          tr,
          circuit,
          "gmep",
          textHeight,
          1.0,
          7,
          "0",
          new Point3d(circuitX, height, 0)
        );
        CreateHorizontalLine(startPoint.X, startPoint.Y, circuit, left, tr, btr);
      }
      CreateBreakerLine(new Point3d(startPoint.X, startPoint.Y - 0.0936, 0), i, left, tr, btr, 2);
    }

    private void CreateHalfSplit2PoleBreakers2Ph(
      Transaction tr,
      BlockTableRecord btr,
      Point3d startPoint,
      Dictionary<string, object> panelData,
      Dictionary<string, double> data,
      bool left,
      bool is2Pole,
      int i
    )
    {
      var (
        descriptions,
        breakers,
        circuits,
        phaseA,
        phaseB,
        phaseC,
        descriptionHighlights,
        descriptionTags
      ) = GetCorrectBreakerData(panelData, left, is2Pole);

      double descriptionX = GetDescriptionX2P(startPoint, left);
      double breakerX = GetBreakerX(startPoint, left);
      double circuitX = GetCircuitX(startPoint, left);
      double textHeight = 0.0725;
      double length = 0.14;

      for (var j = i; j <= i + 3; j++)
      {
        List<string> phaseList = GetPhaseList2P(j, phaseA, phaseB);
        if (Regex.IsMatch(descriptions[j], @"^[a-fA-F0-9]{8}(-[a-fA-F0-9]{4}){3}-[a-fA-F0-9]{12}$"))
        {
          List<Dictionary<string, object>> panelStorage = myForm.RetrieveSavedPanelData();
          foreach (Dictionary<string, object> panel in panelStorage)
          {
            if ((panel["id"] as string).ToLower() == descriptions[j].ToLower())
            {
              descriptions[j] = "PANEL " + panel["panel"] as string;
            }
          }
        }
        string description =
          (descriptionHighlights[j] && descriptions[j] != "EXISTING LOAD")
            ? "(E)" + descriptions[j]
            : descriptions[j];
        string breaker = breakers[j];
        string phase = phaseList[j];
        string circuit = circuits[j];
        double height = startPoint.Y + (-0.890211813771344 - ((j / 2) * 0.1872));
        double phaseX = GetPhaseX2P(j, startPoint, left);

        if (j == i + 2)
        {
          description = "---";
          breakerX += 0.16;
        }
        if (j == i + 3)
        {
          breakerX -= 0.16 + 0.045;
        }

        double yShift = 0;
        if ((j - i) % 2 == 0)
        {
          yShift = 0.0575;
        }
        else
        {
          yShift = -0.0378;
        }
        height = height + yShift;

        if ((j - i) == 0 || (j - i) == 3)
        {
          breaker = breaker + "-1";
          length = 0.23;
        }
        else if ((j - i) == 1)
        {
          length = 0.105;
        }
        else
        {
          length = 0.045;
          breakerX += 0.045;
        }

        CreateAndPositionText(
          tr,
          description.Trim(),
          "gmep",
          textHeight,
          1,
          2,
          "0",
          new Point3d(descriptionX, height, 0)
        );
        if (phase != "0")
        {
          CreateAndPositionCenteredText(
            tr,
            phase,
            "gmep",
            textHeight,
            1.0,
            2,
            "0",
            new Point3d(phaseX, height, 0)
          );
        }
        if (breaker != "")
        {
          CreateAndPositionFittedText(
            tr,
            breaker,
            "gmep",
            textHeight,
            1.0,
            2,
            "0",
            new Point3d(breakerX, height, 0),
            length
          );
        }
        CreateAndPositionText(
          tr,
          circuit,
          "gmep",
          textHeight,
          1.0,
          7,
          "0",
          new Point3d(circuitX, height, 0)
        );
        CreateHorizontalLine(startPoint.X, startPoint.Y, circuit, left, tr, btr);
      }
      CreateBreakerLine(new Point3d(startPoint.X, startPoint.Y - 0.0936, 0), i, left, tr, btr, 2);
    }

    private void Create2PoleBreaker2P(
      Transaction tr,
      BlockTableRecord btr,
      Point3d startPoint,
      Dictionary<string, object> panelData,
      Dictionary<string, double> data,
      bool left,
      bool is2Pole,
      int i
    )
    {
      var (
        descriptions,
        breakers,
        circuits,
        phaseA,
        phaseB,
        phaseC,
        descriptionHighlights,
        descriptionTags
      ) = GetCorrectBreakerData(panelData, left, is2Pole);

      var (_, _, ed) = GetGlobals();
      double descriptionX = GetDescriptionX2P(startPoint, left);
      double breakerX = GetBreakerX(startPoint, left);
      double circuitX = GetCircuitX(startPoint, left);
      double length = 0.14;

      for (var j = i; j <= i + 2; j += 2)
      {
        List<string> phaseList = GetPhaseList2P(j, phaseA, phaseB);
        if (Regex.IsMatch(descriptions[j], @"^[a-fA-F0-9]{8}(-[a-fA-F0-9]{4}){3}-[a-fA-F0-9]{12}$"))
        {
          List<Dictionary<string, object>> panelStorage = myForm.RetrieveSavedPanelData();
          foreach (Dictionary<string, object> panel in panelStorage)
          {
            if ((panel["id"] as string).ToLower() == descriptions[j].ToLower())
            {
              descriptions[j] = "PANEL " + panel["panel"] as string;
            }
          }
        }
        string description =
          (descriptionHighlights[j] && descriptions[j] != "EXISTING LOAD")
            ? "(E)" + descriptions[j]
            : descriptions[j];
        string breaker = breakers[j];
        string phase = phaseList[j];
        string circuit = circuits[j];
        double height = startPoint.Y + (-0.890211813771344 - ((j / 2) * 0.1872));
        double phaseX = GetPhaseX2P(j, startPoint, left);

        if (j == i + 2)
        {
          description = "---";
          breakerX += 0.16;
          length = 0.07;
        }

        CreateAndPositionText(
          tr,
          description.Trim(),
          "gmep",
          0.09375,
          GetDescriptionWidthFactor(description),
          2,
          "0",
          new Point3d(descriptionX, height, 0)
        );
        if (phase != "0")
          CreateAndPositionCenteredText(
            tr,
            phase,
            "gmep",
            0.09375,
            1.0,
            2,
            "0",
            new Point3d(phaseX, height, 0)
          );
        if (breaker != "")
          CreateAndPositionFittedText(
            tr,
            breaker,
            "gmep",
            0.09375,
            1.0,
            2,
            "0",
            new Point3d(breakerX, height, 0),
            length
          );
        CreateAndPositionText(
          tr,
          circuit,
          "gmep",
          0.09375,
          1.0,
          7,
          "0",
          new Point3d(circuitX, height, 0)
        );
        CreateHorizontalLine(startPoint.X, startPoint.Y, circuit, left, tr, btr);
      }

      CreateBreakerLine(startPoint, i, left, tr, btr, 4);
    }

    private void Create3PoleBreaker(
      Transaction tr,
      BlockTableRecord btr,
      Point3d startPoint,
      Dictionary<string, object> panelData,
      Dictionary<string, double> data,
      bool left,
      bool is2Pole,
      int i
    )
    {
      var (
        descriptions,
        breakers,
        circuits,
        phaseA,
        phaseB,
        phaseC,
        descriptionHighlights,
        descriptionTags
      ) = GetCorrectBreakerData(panelData, left, is2Pole);

      double descriptionX = GetDescriptionX(startPoint, left);
      double breakerX = GetBreakerX(startPoint, left);
      double circuitX = GetCircuitX(startPoint, left);
      double length = 0.14;

      for (var j = i; j <= i + 4; j += 2)
      {
        List<string> phaseList = GetPhaseList(j, phaseA, phaseB, phaseC);
        if (Regex.IsMatch(descriptions[j], @"^[a-fA-F0-9]{8}(-[a-fA-F0-9]{4}){3}-[a-fA-F0-9]{12}$"))
        {
          List<Dictionary<string, object>> panelStorage = myForm.RetrieveSavedPanelData();
          foreach (Dictionary<string, object> panel in panelStorage)
          {
            if ((panel["id"] as string).ToLower() == descriptions[j].ToLower())
            {
              descriptions[j] = "PANEL " + panel["panel"] as string;
            }
          }
        }
        string description =
          (descriptionHighlights[j] && descriptions[j] != "EXISTING LOAD")
            ? "(E)" + descriptions[j]
            : descriptions[j];
        string breaker = breakers[j];
        string phase = phaseList[j];
        string circuit = circuits[j];
        double height = startPoint.Y + (-0.890211813771344 - ((j / 2) * 0.1872));
        double phaseX = GetPhaseX(j, startPoint, left);

        if (j == i + 2)
        {
          description = "---";
        }
        else if (j == i + 4)
        {
          description = "---";
          breakerX += 0.16;
          length = 0.07;
        }

        CreateAndPositionText(
          tr,
          description.Trim(),
          "gmep",
          0.09375,
          GetDescriptionWidthFactor(description),
          2,
          "0",
          new Point3d(descriptionX, height, 0)
        );
        if (phase != "0")
          CreateAndPositionCenteredText(
            tr,
            phase,
            "gmep",
            0.09375,
            1.0,
            2,
            "0",
            new Point3d(phaseX, height, 0)
          );
        if (j != i + 2)
          CreateAndPositionFittedText(
            tr,
            breaker,
            "gmep",
            0.09375,
            1.0,
            2,
            "0",
            new Point3d(breakerX, height, 0),
            length
          );
        CreateAndPositionText(
          tr,
          circuit,
          "gmep",
          0.09375,
          1.0,
          7,
          "0",
          new Point3d(circuitX, height, 0)
        );
        CreateHorizontalLine(startPoint.X, startPoint.Y, circuit, left, tr, btr);
      }

      CreateBreakerLine(startPoint, i, left, tr, btr, 6);
    }

    private void CreateBreakerLine(
      Point3d startPoint,
      int i,
      bool left,
      Transaction tr,
      BlockTableRecord btr,
      int span
    )
    {
      double x1,
        x2;
      double height = startPoint.Y + (-0.7488 - (((i + span) / 2) * 0.1872));
      double y1 = height;
      double y2 = height + (span / 2) * 0.1872;

      if (left)
      {
        x1 = startPoint.X + 3.588;
        x2 = startPoint.X + 3.9;
      }
      else
      {
        x1 = startPoint.X + 5.0856;
        x2 = startPoint.X + 5.3976;
      }

      var lineStart = new Point3d(x1, y1, 0);
      var lineEnd = new Point3d(x2, y2, 0);
      var line = new Line(lineStart, lineEnd);
      line.Layer = "0";
      line.ColorIndex = 2;

      btr.AppendEntity(line);
      tr.AddNewlyCreatedDBObject(line, true);
    }

    private double GetCircuitX(Point3d startPoint, bool left)
    {
      if (left)
      {
        return startPoint.X + 3.93681721750636;
      }
      else
      {
        return startPoint.X + 4.87281721750651;
      }
    }

    private double GetBreakerX(Point3d startPoint, bool left)
    {
      if (left)
      {
        return startPoint.X + 3.60379818231218;
      }
      else
      {
        return startPoint.X + 5.10947444486385;
      }
    }

    private double GetDescriptionX(Point3d startPoint, bool left)
    {
      if (left)
      {
        return startPoint.X + 0.063560431161136;
      }
      else
      {
        return startPoint.X + 7.43528640590171;
      }
    }

    private double GetDescriptionX2P(Point3d startPoint, bool left)
    {
      if (left)
      {
        return startPoint.X + 0.0536663060360638;
      }
      else
      {
        return startPoint.X + 7.40509162108179;
      }
    }

    public double GetPhaseX(int i, Point3d startPoint, bool left)
    {
      if (left)
      {
        if (i % 6 == 0 || i % 6 == 1)
        {
          return startPoint.X + 1.64526228334811;
        }
        else if (i % 6 == 2 || i % 6 == 3)
        {
          return startPoint.X + 2.0792421731542;
        }
        else
        {
          return startPoint.X + 2.50445478897294;
        }
      }
      else
      {
        if (i % 6 == 0 || i % 6 == 1)
        {
          return startPoint.X + 6.11211889838299;
        }
        else if (i % 6 == 2 || i % 6 == 3)
        {
          return startPoint.X + 6.53328984899773;
        }
        else
        {
          return startPoint.X + 6.96804695722213;
        }
      }
    }

    public double GetPhaseX2P(int i, Point3d startPoint, bool left)
    {
      if (left)
      {
        if (i % 4 == 0 || i % 4 == 1)
        {
          return startPoint.X + 1.8390082793234;
        }
        else
        {
          return startPoint.X + 2.39546408826883;
        }
      }
      else
      {
        if (i % 4 == 0 || i % 4 == 1)
        {
          return startPoint.X + 6.21960728338948;
        }
        else
        {
          return startPoint.X + 6.83021158846114;
        }
      }
    }

    public List<string> GetPhaseList(
      int i,
      List<string> phaseA,
      List<string> phaseB,
      List<string> phaseC
    )
    {
      if (i % 6 == 0 || i % 6 == 1)
      {
        return phaseA;
      }
      else if (i % 6 == 2 || i % 3 == 3)
      {
        return phaseB;
      }
      else
      {
        return phaseC;
      }
    }

    public List<string> GetPhaseList2P(int i, List<string> phaseA, List<string> phaseB)
    {
      if (i % 4 == 0 || i % 4 == 1)
      {
        return phaseA;
      }
      else
      {
        return phaseB;
      }
    }

    private enum BreakerType
    {
      OutOfRange,
      SplitMini,
      Mini,
      ThreePole,
      TwoPole,
      SinglePole,
    }

    private BreakerType GetBreakerType(List<string> breakers, List<string> circuits, int i)
    {
      // If index i is out of range for the circuits list, return 0.0
      if (i >= circuits.Count)
        return BreakerType.OutOfRange;

      if (
        (i + 2) < breakers.Count
        && breakers[i + 2] == "2"
        && (circuits[i].Contains('A') || circuits[i].Contains('B'))
      )
      {
        return BreakerType.SplitMini;
      }
      // Check if circuit at index i contains 'A' or 'B'
      if (circuits[i].Contains('A') || circuits[i].Contains('B'))
      {
        return BreakerType.Mini;
      }
      // Check if breakers has a value at [i+2] and if it is '3'
      if ((i + 4) < breakers.Count && breakers[i + 4] == "3")
      {
        return BreakerType.ThreePole;
      }
      // Check if breakers has a value at [i+1] and if it is '2'
      if ((i + 2) < breakers.Count && breakers[i + 2] == "2")
      {
        return BreakerType.TwoPole;
      }
      return BreakerType.SinglePole;
    }

    private void ProcessLineData(
      Transaction tr,
      BlockTableRecord btr,
      Point3d startPoint,
      Point3d endPoint,
      double endOfDataY,
      bool is2Pole
    )
    {
      string layerName = "0";
      double left1,
        left2,
        left3,
        right1,
        right2,
        right3;

      if (is2Pole)
      {
        left1 = 1.7549;
        left2 = 2.3023;
        left3 = 0;
        right1 = 6.7259;
        right2 = 0;
        right3 = 7.3632;
        /*
        CreateLine(
          tr,
          btr,
          startPoint.X + 1.85445441972615,
          startPoint.Y - 0.728330362273937,
          startPoint.X + 1.96793252133921,
          startPoint.Y - 0.595665451113291,
          "0"
        );
        CreateLine(
          tr,
          btr,
          startPoint.X + 2.48147136651869,
          startPoint.Y - 0.728330362273937,
          startPoint.X + 2.5949494681322,
          startPoint.Y - 0.595665451113291,
          "0"
        );
        CreateLine(
          tr,
          btr,
          startPoint.X + 4.16751445704995,
          startPoint.Y - 0.728330362273937,
          startPoint.X + 4.280992558663,
          startPoint.Y - 0.595665451113291,
          "0"
        );
        CreateLine(
          tr,
          btr,
          startPoint.X + 4.57272120528569,
          startPoint.Y - 0.728330362273937,
          startPoint.X + 4.68619930689874,
          startPoint.Y - 0.595665451113291,
          "0"
        );
        CreateLine(
          tr,
          btr,
          startPoint.X + 6.22789412120846,
          startPoint.Y - 0.728330362273937,
          startPoint.X + 6.34137222282197,
          startPoint.Y - 0.595665451113291,
          "0"
        );
        CreateLine(
          tr,
          btr,
          startPoint.X + 6.8940934228981,
          startPoint.Y - 0.728330362273937,
          startPoint.X + 7.0075715245116,
          startPoint.Y - 0.595665451113291,
          "0"
        );
        */
      }
      else
      {
        left1 = 1.6224;
        left2 = 2.0488;
        left3 = 2.4752;
        right1 = 6.5104;
        right2 = 6.9368;
        right3 = 7.3632;
        /*
        CreateLine(
          tr,
          btr,
          startPoint.X + 1.8219640114711,
          startPoint.Y + -0.587355127494504,
          startPoint.X + 1.75209866364992,
          startPoint.Y + -0.732578250164607,
          "0"
        );
        CreateLine(
          tr,
          btr,
          startPoint.X + 2.2392415498706,
          startPoint.Y + -0.587355127494504,
          startPoint.X + 2.16937620204942,
          startPoint.Y + -0.732578250164607,
          "0"
        );
        CreateLine(
          tr,
          btr,
          startPoint.X + 2.64064439053459,
          startPoint.Y + -0.587355127494504,
          startPoint.X + 2.57077904271353,
          startPoint.Y + -0.732578250164607,
          "0"
        );
        CreateLine(
          tr,
          btr,
          startPoint.X + 4.28558110707343,
          startPoint.Y + -0.581743684459752,
          startPoint.X + 4.21812491713047,
          startPoint.Y + -0.728299423953047,
          "0"
        );
        CreateLine(
          tr,
          btr,
          startPoint.X + 4.4919520727949,
          startPoint.Y + -0.581743684459752,
          startPoint.X + 4.42449588285183,
          startPoint.Y + -0.728299423953047,
          "0"
        );
        CreateLine(
          tr,
          btr,
          startPoint.X + 4.69832301754843,
          startPoint.Y + -0.581743684459752,
          startPoint.X + 4.63086682760547,
          startPoint.Y + -0.728299423953047,
          "0"
        );
        CreateLine(
          tr,
          btr,
          startPoint.X + 6.28764850398056,
          startPoint.Y + -0.586040159740406,
          startPoint.X + 6.21478297900239,
          startPoint.Y + -0.730701330926394,
          "0"
        );
        CreateLine(
          tr,
          btr,
          startPoint.X + 6.69812260049363,
          startPoint.Y + -0.586040159740406,
          startPoint.X + 6.62525707551547,
          startPoint.Y + -0.730701330926394,
          "0"
        );
        CreateLine(
          tr,
          btr,
          startPoint.X + 7.10859666269549,
          startPoint.Y + -0.586040159740406,
          startPoint.X + 7.03573113771732,
          startPoint.Y + -0.730701330926394,
          "0"
        );
        */
      }

      var linesData = new (double[] distances, double startY, double endY, string layer)[]
      {
        (new double[] { 2.2222, 5.0666, 6.9368 }, startPoint.Y, startPoint.Y - 0.3744, layerName),
        (
          new double[]
          {
            left1,
            2.9016,
            3.1304,
            3.3592,
            3.5880,
            3.9000,
            4.1496,
            4.8360,
            5.0856,
            5.3976,
            5.6264,
            5.8552,
            6.0840,
            right3,
          },
          startPoint.Y - 0.3744,
          endOfDataY,
          layerName
        ),
        (
          new double[] { left2, left3, right1, right2 },
          startPoint.Y - 0.3744 - (0.3744 / 2),
          endOfDataY,
          layerName
        ),
      };

      foreach (var lineData in linesData)
      {
        CreateVerticalLines(
          tr,
          btr,
          startPoint,
          lineData.distances,
          lineData.startY,
          lineData.endY,
          lineData.layer
        );
      }

      var linesData2 = new (double startX, double startY, double endX, double endY, string layer)[]
      {
        (startPoint.X, startPoint.Y - 0.3744, endPoint.X, startPoint.Y - 0.3744, layerName),
        (startPoint.X, startPoint.Y - 0.7488, endPoint.X, startPoint.Y - 0.7488, layerName),
        (
          startPoint.X + 2.2222,
          startPoint.Y - (0.3744 / 2),
          endPoint.X,
          startPoint.Y - (0.3744 / 2),
          layerName
        ),
        (
          startPoint.X + left1,
          startPoint.Y - (0.3744 / 2) - 0.3744,
          startPoint.X + 2.9016,
          startPoint.Y - (0.3744 / 2) - 0.3744,
          layerName
        ),
        (
          startPoint.X + 6.0840,
          startPoint.Y - (0.3744 / 2) - 0.3744,
          startPoint.X + 7.3632,
          startPoint.Y - (0.3744 / 2) - 0.3744,
          layerName
        ),
        (
          startPoint.X + 4.1496,
          startPoint.Y - (0.3744 / 2) - 0.3744,
          startPoint.X + 4.8360,
          startPoint.Y - (0.3744 / 2) - 0.3744,
          layerName
        ),
        //(
        //  startPoint.X + 8.28490642235897,
        //  startPoint.Y + -0.153581773169606,
        //  startPoint.X + 8.37682368466574,
        //  startPoint.Y + -0.0461231951291552,
        //  "0"
        //),
      };

      CreateLines(tr, btr, linesData2);
    }

    public static ObjectId GetTextStyleId(string styleName)
    {
      var (doc, db, _) = PanelCommands.GetGlobals();
      var textStyleTable = (TextStyleTable)db.TextStyleTableId.GetObject(OpenMode.ForRead);

      if (textStyleTable.Has(styleName))
      {
        return textStyleTable[styleName];
      }
      else
      {
        // Return the ObjectId of the "gmep" style
        return textStyleTable["gmep"];
      }
    }

    private double CreateHorizontalLine(
      double startPointX,
      double startPointY,
      string circuitNumber,
      bool left,
      Transaction tr,
      BlockTableRecord btr
    )
    {
      int circuitNumReducer;
      double lineStartX;
      double lineStartX2;
      int circuitNumInt;
      double deltaY = 0.187200000000021; // Change this value if needed

      if (left)
      {
        circuitNumReducer = 1;
        lineStartX = startPointX;
        lineStartX2 = startPointX + 4.14960000000019;
      }
      else
      {
        circuitNumReducer = 2;
        lineStartX = startPointX + 4.8360;
        lineStartX2 = startPointX + 8.9856;
      }

      if (circuitNumber.Contains('A') || circuitNumber.Contains('B'))
      {
        // Remove 'A' or 'B' from the string
        if (left)
        {
          circuitNumInt = circuitNumLeft;
          if (circuitNumber.Contains('B'))
          {
            circuitNumLeft += 2;
          }
        }
        else
        {
          circuitNumInt = circuitNumRight;
          if (circuitNumber.Contains('B'))
          {
            circuitNumRight += 2;
          }
        }
        CreateCircuitLine(
          circuitNumInt,
          circuitNumReducer,
          startPointY,
          deltaY,
          lineStartX,
          lineStartX2,
          tr,
          btr,
          true
        );
      }
      else
      {
        if (left)
        {
          circuitNumInt = circuitNumLeft;
          circuitNumLeft += 2;
        }
        else
        {
          circuitNumInt = circuitNumRight;
          circuitNumRight += 2;
        }
      }

      return CreateCircuitLine(
        circuitNumInt,
        circuitNumReducer,
        startPointY,
        deltaY,
        lineStartX,
        lineStartX2,
        tr,
        btr
      );
    }

    private double CreateCircuitLine(
      int circuitNumInt,
      int circuitNumReducer,
      double startPointY,
      double deltaY,
      double lineStartX,
      double lineStartX2,
      Transaction tr,
      BlockTableRecord btr,
      bool half = false
    )
    {
      circuitNumInt = (circuitNumInt - circuitNumReducer) / 2;
      double lineStartY = startPointY - (0.935999999999979 + (deltaY * circuitNumInt));
      if (half)
        lineStartY += deltaY / 2;
      double lineEndY = lineStartY;

      var lineStart = new Point3d(lineStartX, lineStartY, 0);
      var lineEnd = new Point3d(lineStartX2, lineEndY, 0);
      var line = new Line(lineStart, lineEnd) { Layer = "0" };

      btr.AppendEntity(line);
      tr.AddNewlyCreatedDBObject(line, true);

      return line.StartPoint.Y;
    }

    private double GetDescriptionWidthFactor(string description)
    {
      description = description.Trim();
      if (description.Length <= 20)
      {
        return 0.8;
      }
      if (description.Length <= 24)
      {
        return 0.75;
      }
      if (description.Length <= 28)
      {
        return 0.65;
      }
      return 0.6;
    }
  }
}
