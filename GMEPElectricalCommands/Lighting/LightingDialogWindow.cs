using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using ElectricalCommands.ElectricalEntity;
using GMEPElectricalCommands.GmepDatabase;

namespace ElectricalCommands.Lighting
{
  public partial class LightingDialogWindow : Form
  {
    private List<LightingFixture> lightingFixtureList;
    private List<ListViewItem> lightingFixtureListViewList;
    private List<LightingControl> lightingControlList;
    private List<ListViewItem> lightingControlListViewList;
    private List<LightingSignage> lightingSignageList;
    private List<ListViewItem> lightingSignageListViewList;
    private string projectId;
    public GmepDatabase gmepDb = new GmepDatabase();
    private List<ElectricalEntity.Panel> panelList;
    private bool isLoading;

    public LightingDialogWindow()
    {
      InitializeComponent();
    }

    public void InitializeModal()
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Core
        .Application
        .DocumentManager
        .MdiActiveDocument;
      projectId = gmepDb.GetProjectId(CADObjectCommands.GetProjectNoFromFileName());
      panelList = gmepDb.GetPanels(projectId);
      lightingFixtureList = gmepDb.GetLightingFixtures(projectId);
      lightingControlList = gmepDb.GetLightingControls(projectId);
      lightingSignageList = gmepDb.GetLightingSignage(projectId);
      Console.WriteLine("lightingFixtureList: "+ lightingFixtureList.Count);
      CreateLightingFixtureListView();
      CreateLightingControlListView();
      CreateLightingSignageListView();

      isLoading = false;
    }

    public static Dictionary<string, int> GetNumObjectsOnPlan(string propName)
    {
      Dictionary<string, int> objDict = new Dictionary<string, int>();
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;

      Database db = doc.Database;
      Editor ed = doc.Editor;
      Transaction tr = db.TransactionManager.StartTransaction();
      using (tr)
      {
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        var modelSpace = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
        foreach (ObjectId id in modelSpace)
        {
          try
          {
            BlockReference br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
            if (br != null && br.IsDynamicBlock)
            {
              DynamicBlockReferencePropertyCollection pc =
                br.DynamicBlockReferencePropertyCollection;
              PlaceableElectricalEntity eq = new PlaceableElectricalEntity();
              foreach (DynamicBlockReferenceProperty prop in pc)
              {
                if (prop.PropertyName == propName)
                {
                  if (!objDict.ContainsKey(prop.Value as string))
                  {
                    objDict[prop.Value as string] = 1;
                  }
                  else
                  {
                    objDict[prop.Value as string] += 1;
                  }
                }
              }
            }
          }
          catch (Exception e) { }
        }
      }
      return objDict;
    }

    private void CreateLightingSignageListView(bool updateOnly = false) {
      if (updateOnly) {
        LightingSignageListView.Items.Clear();
      }
      LightingSignageListView.View = View.Details;
      LightingSignageListView.FullRowSelect = true;
      Dictionary<string, int> controlDict = GetNumObjectsOnPlan("gmep_lighting_signage_id");
      foreach (LightingSignage signage in lightingSignageList) {
        int placed = 0;
        if (!string.IsNullOrEmpty(signage.Id) && controlDict.ContainsKey(signage.Id)) {
          placed = controlDict[signage.Id];
        }
        ListViewItem item = new ListViewItem(signage.Tag, 0);
        item.SubItems.Add(signage.Volt.ToString());
        item.SubItems.Add(signage.Description);
        item.SubItems.Add(signage.IndoorOutdoor);
        LightingSignageListView.Items.Add(item);
      }
      if (!updateOnly) {
        LightingSignageListView.Columns.Add("Tag", -2, HorizontalAlignment.Left);
        LightingSignageListView.Columns.Add("Volt", -2, HorizontalAlignment.Left);
        LightingSignageListView.Columns.Add("Description", -2, HorizontalAlignment.Left);
        LightingSignageListView.Columns.Add("Indoor/Outdoor", -2, HorizontalAlignment.Left);
      }
      Console.WriteLine($"projectId Name: {projectId}");
    }
    private void CreateLightingControlListView(bool updateOnly = false)
    {
      if (updateOnly)
      {
        LightingControlsListView.Items.Clear();
      }
      LightingControlsListView.View = View.Details;
      LightingControlsListView.FullRowSelect = true;
      Dictionary<string, int> controlDict = GetNumObjectsOnPlan("gmep_lighting_control_id");
      foreach (LightingControl control in lightingControlList)
      {
        int placed = 0;
        if (controlDict.ContainsKey(control.Id))
        {
          placed = controlDict[control.Id];
        }
        ListViewItem item = new ListViewItem(control.Name, 0);
        item.SubItems.Add(control.ControlType);
        item.SubItems.Add(control.HasOccupancy ? "Yes" : "No");
        item.SubItems.Add(placed.ToString());
        item.SubItems.Add(control.Id);
        LightingControlsListView.Items.Add(item);
      }
      if (!updateOnly)
      {
        LightingControlsListView.Columns.Add("Tag", -2, HorizontalAlignment.Left);
        LightingControlsListView.Columns.Add("Control Type", -2, HorizontalAlignment.Left);
        LightingControlsListView.Columns.Add("Occupancy", -2, HorizontalAlignment.Left);
        LightingControlsListView.Columns.Add("Placed", -2, HorizontalAlignment.Left);
      }
    }

    private void CreateLightingFixtureListView(bool updateOnly = false)
    {
      if (updateOnly)
      {
        LightingFixturesListView.Items.Clear();
      }
      LightingFixturesListView.View = View.Details;
      LightingFixturesListView.FullRowSelect = true;
      Dictionary<string, int> fixtureDict = GetNumObjectsOnPlan("gmep_lighting_fixture_id");
      foreach (LightingFixture fixture in lightingFixtureList)
      {
        Console.WriteLine("lightingFixtureList: " + fixture.Name);
        int placed = 0;
        if (fixtureDict.ContainsKey(fixture.Id))
        {
          placed = fixtureDict[fixture.Id];
        }
        ListViewItem item = new ListViewItem(fixture.Name, 0);
        item.SubItems.Add(fixture.BlockName);
        item.SubItems.Add(fixture.Voltage.ToString());
        item.SubItems.Add(fixture.Qty.ToString());
        item.SubItems.Add(placed.ToString());
        item.SubItems.Add(fixture.Mounting);
        item.SubItems.Add(fixture.Description);
        item.SubItems.Add(fixture.Manufacturer);
        item.SubItems.Add(fixture.ModelNo);
        item.SubItems.Add("LED");
        item.SubItems.Add(Math.Round(fixture.Wattage, 1).ToString());
        item.SubItems.Add(fixture.Notes);
        item.SubItems.Add(fixture.Id);
        LightingFixturesListView.Items.Add(item);
      }
      if (!updateOnly)
      {
        LightingFixturesListView.Columns.Add("Tag", -2, HorizontalAlignment.Left);
        LightingFixturesListView.Columns.Add("Legend", -2, HorizontalAlignment.Left);
        LightingFixturesListView.Columns.Add("Volt", -2, HorizontalAlignment.Left);
        LightingFixturesListView.Columns.Add("Count", -2, HorizontalAlignment.Left);
        LightingFixturesListView.Columns.Add("Placed", -2, HorizontalAlignment.Left);
        LightingFixturesListView.Columns.Add("Mount", -2, HorizontalAlignment.Left);
        LightingFixturesListView.Columns.Add("Description", -2, HorizontalAlignment.Left);
        LightingFixturesListView.Columns.Add("Manufacturer", -2, HorizontalAlignment.Left);
        LightingFixturesListView.Columns.Add("Model Number", -2, HorizontalAlignment.Left);
        LightingFixturesListView.Columns.Add("Lamps", -2, HorizontalAlignment.Left);
        LightingFixturesListView.Columns.Add("Input Watts", -2, HorizontalAlignment.Left);
        LightingFixturesListView.Columns.Add("Notes", -2, HorizontalAlignment.Left);
      }
    }

    private void CreateLightingFixtureSchedule(
      Document doc,
      Database db,
      Editor ed,
      Point3d startPoint
    )
    {
      var spaceId =
        (db.TileMode == true)
          ? SymbolUtilityServices.GetBlockModelSpaceId(db)
          : SymbolUtilityServices.GetBlockPaperSpaceId(db);
      List<LightingLocation> lightingLocations = gmepDb.GetLightingLocations(projectId);
      
      if (lightingLocations == null || lightingLocations.Count == 0)
      {
        lightingLocations = new List<LightingLocation>();
        lightingLocations.Add(new LightingLocation("", "", false, ""));
      }
      foreach (LightingLocation lightingLocation in lightingLocations)
      {
        using (var tr = db.TransactionManager.StartTransaction())
        {
          var btr = (BlockTableRecord)tr.GetObject(spaceId, OpenMode.ForWrite);
          Table tb = new Table();
          tb.TableStyle = db.Tablestyle;
          tb.Position = startPoint;
          List<LightingFixture> fixtures = lightingFixtureList.FindAll(f =>
            f.LocationId == lightingLocation.Id
          );

          int tableRows = fixtures.Count + 3;
          int tableCols = 11;
          tb.SetSize(tableRows, tableCols);
          tb.SetRowHeight(0.8911);
          tb.Cells[0, 0].TextString =
            "LIGHTING FIXTURE SCHEDULE"
            + (
              !String.IsNullOrEmpty(lightingLocation.LocationName)
                ? " - " + lightingLocation.LocationName.ToUpper()
                : ""
            );
          tb.Rows[0].Height = 0.6117;
          tb.Rows[1].Height = 0.5357;
          tb.Rows[tableRows - 1].Height = 0.7023;
          var textStyleId = PanelCommands.GetTextStyleId("gmep");
          var titleStyleId = PanelCommands.GetTextStyleId("section title");
          CellRange range = CellRange.Create(tb, tableRows - 1, 0, tableRows - 1, tableCols - 1);
          tb.MergeCells(range);
          tb.Layer = "E-TXT1";
          tb.Cells[1, 0].TextString = "MARK";
          tb.Cells[1, 1].TextString = "LEGEND";
          tb.Cells[1, 2].TextString = "VOLT";
          tb.Cells[1, 3].TextString = "COUNT";
          tb.Cells[1, 4].TextString = "MOUNT";
          tb.Cells[1, 5].TextString = "DESCRIPTION";
          tb.Cells[1, 6].TextString = "MANUFACTURER";
          tb.Cells[1, 7].TextString = "MODEL NUMBER";
          tb.Cells[1, 8].TextString = "LAMPS";
          tb.Cells[1, 9].TextString = "INPUT WATTS";
          tb.Cells[1, 10].TextString = "NOTES";
          tb.Columns[0].Width = 0.8395;
          tb.Columns[1].Width = 0.7481;
          tb.Columns[2].Width = 0.7157;
          tb.Columns[3].Width = 0.6763;
          tb.Columns[4].Width = 1.0107;
          tb.Columns[5].Width = 3.0413;
          tb.Columns[6].Width = 1.4384;
          tb.Columns[7].Width = 1.5674;
          tb.Columns[8].Width = 0.7886;
          tb.Columns[9].Width = 0.7757;
          tb.Columns[10].Width = 1.3997;
          for (int i = 0; i < tableRows; i++)
          {
            for (int j = 0; j < tableCols; j++)
            {
              if (i <= 1)
              {
                tb.Cells[i, j].TextStyleId = titleStyleId;
                if (i == 0)
                {
                  tb.Cells[i, j].TextHeight = (0.25);
                }
                else
                {
                  tb.Cells[i, j].TextHeight = (0.125);
                }
              }
              else
              {
                if (j == 5 || j == 7 || j == 10)
                {
                  tb.Cells[i, j].Alignment = CellAlignment.MiddleLeft;
                }
                else
                {
                  tb.Cells[i, j].Alignment = CellAlignment.MiddleCenter;
                }
                tb.Cells[i, j].TextStyleId = textStyleId;
                tb.Cells[i, j].TextHeight = (0.0938);
              }
              if (i == tableRows - 1)
              {
                tb.Cells[i, j].Alignment = CellAlignment.TopLeft;
              }
            }
          }
          for (int i = 0; i < fixtures.Count; i++)
          {
            int row = i + 2;
            tb.Cells[row, 2].TextString = fixtures[i].Voltage.ToString();
            tb.Cells[row, 3].TextString = fixtures[i].Qty.ToString();
            tb.Cells[row, 4].TextString = fixtures[i].Mounting.ToUpper();
            tb.Cells[row, 5].TextString = fixtures[i].Description.ToUpper();
            tb.Cells[row, 6].TextString = fixtures[i].Manufacturer.ToUpper();
            tb.Cells[row, 7].TextString = fixtures[i].ModelNo.ToUpper();
            tb.Cells[row, 8].TextString = "LED";
            tb.Cells[row, 9].TextString = fixtures[i].Wattage.ToString();
            tb.Cells[row, 10].TextString = fixtures[i].Notes.ToUpper();
          }
          tb.Cells[tableRows - 1, 0].TextString =
            "NOTES:\n  1) VERIFY WITH OWNER OR ARCHITECT BEFORE PURCHASING THE LIGHTING FIXTURES.\n 2) LIGHTING ABOVE FOOD OR UTENSILS SHALL BE SHATTERPROOF.";
          BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
          btr.AppendEntity(tb);
          tr.AddNewlyCreatedDBObject(tb, true);
          int r = 0;
          foreach (LightingFixture fixture in fixtures)
          {
            try
            {
              ObjectId blockId = bt[fixture.BlockName];
              if (fixture.EmCapable)
              {
                using (
                  BlockReference acBlkRef = new BlockReference(
                    new Point3d(startPoint.X + 1.0342, startPoint.Y - (1.58 + (0.8911 * r)), 0),
                    blockId
                  )
                )
                {
                  BlockTableRecord acCurSpaceBlkTblRec;
                  acCurSpaceBlkTblRec =
                    tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                  acCurSpaceBlkTblRec.AppendEntity(acBlkRef);

                  acBlkRef.Layer = "E-SYM1";
                  acBlkRef.ScaleFactors = new Scale3d(fixture.PaperSpaceScale);
                  tr.AddNewlyCreatedDBObject(acBlkRef, true);
                }
                using (
                  BlockReference acBlkRef = new BlockReference(
                    new Point3d(startPoint.X + 1.3892, startPoint.Y - (1.58 + (0.8911 * r)), 0),
                    blockId
                  )
                )
                {
                  BlockTableRecord acCurSpaceBlkTblRec;
                  acCurSpaceBlkTblRec =
                    tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                  acCurSpaceBlkTblRec.AppendEntity(acBlkRef);

                  acBlkRef.Layer = "E-SYM1";
                  acBlkRef.ScaleFactors = new Scale3d(fixture.PaperSpaceScale);
                  tr.AddNewlyCreatedDBObject(acBlkRef, true);
                }
                ObjectId emBlockId = bt[fixture.BlockName + " EM"];
                using (
                  BlockReference acBlkRef = new BlockReference(
                    new Point3d(startPoint.X + 1.3892, startPoint.Y - (1.58 + (0.8911 * r)), 0),
                    emBlockId
                  )
                )
                {
                  BlockTableRecord acCurSpaceBlkTblRec;
                  acCurSpaceBlkTblRec =
                    tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                  acCurSpaceBlkTblRec.AppendEntity(acBlkRef);

                  acBlkRef.Layer = "E-SYM1";
                  acBlkRef.ScaleFactors = new Scale3d(fixture.PaperSpaceScale);
                  tr.AddNewlyCreatedDBObject(acBlkRef, true);
                }
              }
              else
              {
                using (
                  BlockReference acBlkRef = new BlockReference(
                    new Point3d(startPoint.X + 1.2117, startPoint.Y - (1.58 + (0.8911 * r)), 0),
                    blockId
                  )
                )
                {
                  BlockTableRecord acCurSpaceBlkTblRec;
                  acCurSpaceBlkTblRec =
                    tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                  acCurSpaceBlkTblRec.AppendEntity(acBlkRef);

                  acBlkRef.Layer = "E-SYM1";
                  acBlkRef.ScaleFactors = new Scale3d(fixture.PaperSpaceScale);
                  tr.AddNewlyCreatedDBObject(acBlkRef, true);
                }
              }
              ObjectId tagBlockId = bt["4CFM"];
              using (
                BlockReference acBlkRef = new BlockReference(
                  new Point3d(startPoint.X + 0.1415, startPoint.Y - (1.58 + (0.8911 * r)), 0),
                  tagBlockId
                )
              )
              {
                BlockTableRecord acCurSpaceBlkTblRec;
                acCurSpaceBlkTblRec =
                  tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                acCurSpaceBlkTblRec.AppendEntity(acBlkRef);

                acBlkRef.Layer = "E-TXT1";
                tr.AddNewlyCreatedDBObject(acBlkRef, true);
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
                AttributeDefinition attrDef = new AttributeDefinition();
                attrDef.Position = new Point3d(
                  startPoint.X + 0.4025,
                  startPoint.Y - (1.47 + (0.8911 * r)),
                  0
                );

                var textStyle = (TextStyleTableRecord)
                  tr.GetObject(gmepTextStyleId, OpenMode.ForRead);
                double widthFactor = 1;
                if (
                  textStyle.FileName.ToLower().Contains("architxt")
                  || textStyle.FileName.ToLower().Contains("a2")
                )
                {
                  widthFactor = 0.85;
                }

                attrDef.LockPositionInBlock = false;
                attrDef.Tag = "tag";
                attrDef.IsMTextAttributeDefinition = false;
                attrDef.TextString = fixture.Name;
                attrDef.Justify = AttachmentPoint.MiddleCenter;
                attrDef.Visible = true;
                attrDef.Invisible = false;
                attrDef.Constant = false;
                attrDef.Height = 0.0938;
                attrDef.WidthFactor = widthFactor;
                attrDef.TextStyleId = gmepTextStyleId;
                attrDef.Layer = "0";

                AttributeReference attrRef = new AttributeReference();
                attrRef.SetAttributeFromBlock(
                  attrDef,
                  Matrix3d.Displacement(new Vector3d(attrDef.Position.X, attrDef.Position.Y, 0))
                );
                acBlkRef.AttributeCollection.AppendAttribute(attrRef);

                AttributeDefinition attrDef2 = new AttributeDefinition();
                attrDef2.Position = new Point3d(
                  startPoint.X + 0.4025,
                  startPoint.Y - (1.70 + (0.8911 * r)),
                  0
                );
                attrDef2.LockPositionInBlock = false;
                attrDef2.Tag = "wattage";
                attrDef2.IsMTextAttributeDefinition = false;
                attrDef2.TextString = fixture.Wattage.ToString();
                attrDef2.Justify = AttachmentPoint.MiddleCenter;
                attrDef2.Visible = true;
                attrDef2.Invisible = false;
                attrDef2.Constant = false;
                attrDef2.Height = 0.0938;
                attrDef2.WidthFactor = widthFactor;
                attrDef2.TextStyleId = gmepTextStyleId;
                attrDef2.Layer = "0";

                AttributeReference attrRef2 = new AttributeReference();
                attrRef2.SetAttributeFromBlock(
                  attrDef2,
                  Matrix3d.Displacement(new Vector3d(attrDef2.Position.X, attrDef2.Position.Y, 0))
                );
                acBlkRef.AttributeCollection.AppendAttribute(attrRef2);
                acBlkRef.Layer = "E-TXT1";
                tr.AddNewlyCreatedDBObject(acBlkRef, true);
              }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex) { }
            r++;
          }
          tr.Commit();
        }
        startPoint = new Point3d(startPoint.X + 13.1, startPoint.Y, 0);
      }
    }

    private void CreateLightingFixtureScheduleButton_Click(object sender, EventArgs e)
    {
      Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.WindowState = Autodesk
        .AutoCAD
        .Windows
        .Window
        .State
        .Maximized;
      Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Window.Focus();
      Point3d startingPoint;
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      var currentView = ed.GetCurrentView();
      using (
        DocumentLock docLock =
          Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
      )
      {
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          var promptOptions = new PromptPointOptions("\nSelect upper left point:");
          var promptResult = ed.GetPoint(promptOptions);
          currentView = ed.GetCurrentView();
          if (promptResult.Status == PromptStatus.OK)
            startingPoint = promptResult.Value;
          else
          {
            return;
          }
        }
        CreateLightingFixtureSchedule(doc, db, ed, startingPoint);
      }
      ed.SetCurrentView(currentView);
    }

    private void PlaceControl_Click(object sender, EventArgs e)
    {
      if (LightingControlsListView.SelectedItems.Count == 0)
      {
        return;
      }
      using (
        DocumentLock docLock =
          Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
      )
      {
        Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.WindowState = Autodesk
          .AutoCAD
          .Windows
          .Window
          .State
          .Maximized;
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Window.Focus();
        if (CADObjectCommands.Scale == -1.0)
        {
          CADObjectCommands.SetScale();
        }
        Document doc = Autodesk
          .AutoCAD
          .ApplicationServices
          .Core
          .Application
          .DocumentManager
          .MdiActiveDocument;
        Editor ed = doc.Editor;
        Database db = doc.Database;
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          LayerTable acLyrTbl;
          acLyrTbl = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

          string sLayerName = "E-LITE-FIXT";

          if (acLyrTbl.Has(sLayerName) == true)
          {
            db.Clayer = acLyrTbl[sLayerName];
            tr.Commit();
          }
        }
        int numSubitems = LightingControlsListView.SelectedItems[0].SubItems.Count;
        Dictionary<string, int> controlDict = GetNumObjectsOnPlan("gmep_lighting_control_id");
        for (int si = 0; si < LightingControlsListView.SelectedItems.Count; si++)
        {
          foreach (ElectricalEntity.LightingControl control in lightingControlList)
          {
            if (
              control.Id
              != LightingControlsListView.SelectedItems[si].SubItems[numSubitems - 1].Text
            )
            {
              continue;
            }
            int currentNumControls = 0;
            if (controlDict.ContainsKey(control.Id))
            {
              currentNumControls = controlDict[control.Id];
            }
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
            try
            {
              Point3d point;
              double rotation = 0;
              using (Transaction tr = db.TransactionManager.StartTransaction())
              {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr;
                BlockReference br = CADObjectCommands.CreateBlockReference(
                  tr,
                  bt,
                  blockName,
                  out btr,
                  out point
                );

                if (br != null)
                {
                  BlockTableRecord curSpace = (BlockTableRecord)
                    tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                  RotateJig rotateJig = new RotateJig(br);
                  PromptResult rotatePromptResult = ed.Drag(rotateJig);

                  if (rotatePromptResult.Status != PromptStatus.OK)
                  {
                    return;
                  }
                  rotation = br.Rotation;

                  Console.WriteLine(rotation);

                  double xTransform = 0;
                  double yTransform = 0;

                  if (rotation > 4.7 && rotation < 4.72)
                  {
                    yTransform = -4.5 * 0.25 / CADObjectCommands.Scale;
                  }
                  if (rotation > 1.5 && rotation < 1.6)
                  {
                    xTransform = -4.5 * 0.25 / CADObjectCommands.Scale;
                  }
                  if (rotation > 3 && rotation < 3.2)
                  {
                    xTransform = -4.5 * 0.25 / CADObjectCommands.Scale;
                    yTransform = -4.5 * 0.25 / CADObjectCommands.Scale;
                  }

                  Console.WriteLine($"Rotation: {rotation}");
                  Console.WriteLine($"xTransform: {xTransform}");
                  Console.WriteLine($"yTransform: {yTransform}");

                  curSpace.AppendEntity(br);

                  tr.AddNewlyCreatedDBObject(br, true);
                  blockId = br.Id;

                  //Setting Attributes
                  foreach (ObjectId objId in btr)
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
                          Point3d p = attDef.Position.TransformBy(br.BlockTransform);
                          if (rotation > 1.5 && rotation < 1.6)
                          {
                            p = new Point3d(p.X + xTransform, p.Y, 0);
                          }
                          attRef.Position = p;
                          attRef.TextString = control.Name;
                          attRef.Height = 0.0938 / CADObjectCommands.Scale * 12;
                          attRef.WidthFactor = 1;
                          attRef.HorizontalMode = TextHorizontalMode.TextLeft;
                          attRef.VerticalMode = TextVerticalMode.TextVerticalMid;
                          attRef.Justify = AttachmentPoint.BaseLeft;
                          attRef.Rotation = attRef.Rotation - rotation;
                        }
                        if (attRef.Tag == "D")
                        {
                          Point3d dPoint = attDef.Position.TransformBy(br.BlockTransform);
                          Console.WriteLine(dPoint.Y);
                          dPoint = new Point3d(dPoint.X + xTransform, dPoint.Y + yTransform, 0);
                          Console.WriteLine(dPoint.Y);

                          attRef.Position = dPoint;
                          attRef.Height = 0.0938 / CADObjectCommands.Scale * 12;
                          attRef.WidthFactor = 1;
                          attRef.TextString = attRef.Tag;
                          attRef.HorizontalMode = TextHorizontalMode.TextLeft;
                          attRef.VerticalMode = TextVerticalMode.TextVerticalMid;
                          attRef.Justify = AttachmentPoint.BaseLeft;

                          Matrix3d rotationMatrix = Matrix3d.Rotation(
                            -rotation,
                            Vector3d.ZAxis,
                            dPoint
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
                DynamicBlockReferencePropertyCollection pc =
                  br.DynamicBlockReferencePropertyCollection;
                foreach (DynamicBlockReferenceProperty prop in pc)
                {
                  if (
                    prop.PropertyName == "gmep_lighting_control_id"
                    && prop.Value as string == "0"
                  )
                  {
                    prop.Value = control.Id;
                  }
                  if (
                    prop.PropertyName == "gmep_lighting_control_tag"
                    && prop.Value as string == "0"
                  )
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
      CreateLightingFixtureListView(true);
    }

    private void PlaceFixture_Click(object sender, EventArgs e)
    {
      if (LightingFixturesListView.SelectedItems.Count == 0)
      {
        return;
      }
      using (
        DocumentLock docLock =
          Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
      )
      {
        Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.WindowState = Autodesk
          .AutoCAD
          .Windows
          .Window
          .State
          .Maximized;
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Window.Focus();
        if (CADObjectCommands.Scale == -1.0)
        {
          CADObjectCommands.SetScale();
        }
        Document doc = Autodesk
          .AutoCAD
          .ApplicationServices
          .Core
          .Application
          .DocumentManager
          .MdiActiveDocument;
        Editor ed = doc.Editor;
        Database db = doc.Database;
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          LayerTable acLyrTbl;
          acLyrTbl = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;

          string sLayerName = "E-LITE-FIXT";

          if (acLyrTbl.Has(sLayerName) == true)
          {
            db.Clayer = acLyrTbl[sLayerName];
            tr.Commit();
          }
        }
        int numSubitems = LightingFixturesListView.SelectedItems[0].SubItems.Count;
        Dictionary<string, int> fixtureDict = GetNumObjectsOnPlan("gmep_lighting_fixture_id");
        for (int si = 0; si < LightingFixturesListView.SelectedItems.Count; si++)
        {
          foreach (ElectricalEntity.LightingFixture fixture in lightingFixtureList)
          {
            if (
              fixture.Id
              != LightingFixturesListView.SelectedItems[si].SubItems[numSubitems - 1].Text
            )
            {
              continue;
            }
            int currentNumFixtures = 0;
            if (fixtureDict.ContainsKey(fixture.Id))
            {
              currentNumFixtures = fixtureDict[fixture.Id];
            }
            for (int i = currentNumFixtures; i < fixture.Qty; i++)
            {
              ed.WriteMessage(
                "\nPlace "
                  + (i + 1).ToString()
                  + "/"
                  + fixture.Qty.ToString()
                  + " for '"
                  + fixture.Name
                  + "'"
              );
              ObjectId blockId;
              try
              {
                Point3d point;
                double rotation = 0;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                  BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                  BlockTableRecord block;

                  if (String.IsNullOrEmpty(fixture.BlockName))
                  {
                    fixture.BlockName = "GMEP LTG CUSTOM";
                    fixture.LabelTransformVX = -12;
                    fixture.LabelTransformVY = -18;
                    fixture.LabelTransformHX = -24;
                    fixture.LabelTransformHY = -18;
                  }

                  BlockReference br = CADObjectCommands.CreateBlockReference(
                    tr,
                    bt,
                    fixture.BlockName,
                    out block,
                    out point
                  );

                  if (br != null)
                  {
                    BlockTableRecord curSpace = (BlockTableRecord)
                      tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    if (fixture.Rotate)
                    {
                      RotateJig rotateJig = new RotateJig(br);
                      PromptResult rotatePromptResult = ed.Drag(rotateJig);

                      if (rotatePromptResult.Status != PromptStatus.OK)
                      {
                        return;
                      }
                      rotation = br.Rotation;
                    }

                    br.Layer = "E-LITE-EQPM";

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
                          if (attDef.Tag == "LIGHTING_NAME")
                          {
                            attRef.TextString = fixture.Name;
                          }
                          if (attDef.Tag == "LIGHTING_CIRCUIT")
                          {
                            attRef.TextString = "#~";
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
                  DynamicBlockReferencePropertyCollection pc =
                    br.DynamicBlockReferencePropertyCollection;

                  foreach (DynamicBlockReferenceProperty prop in pc)
                  {
                    if (prop.PropertyName == "gmep_lighting_id" && prop.Value as string == "0")
                    {
                      prop.Value = Guid.NewGuid().ToString();
                    }
                    if (
                      prop.PropertyName == "gmep_lighting_fixture_id"
                      && prop.Value as string == "0"
                    )
                    {
                      prop.Value = fixture.Id;
                    }
                    if (prop.PropertyName == "gmep_lighting_name" && prop.Value as string == "0")
                    {
                      prop.Value = fixture.Name;
                    }
                  }
                  tr.Commit();
                }
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                  BlockReference br = (BlockReference)tr.GetObject(blockId, OpenMode.ForWrite);
                  Point3d position = new Point3d(
                    point.X + fixture.LabelTransformVX,
                    point.Y
                      + fixture.LabelTransformVY
                      + (
                        (CADObjectCommands.Scale - 0.25)
                        * 12
                        * Math.Pow(0.25 / CADObjectCommands.Scale, 1.5)
                      ),
                    0
                  );
                  Point3d position2 = new Point3d(
                    point.X + fixture.LabelTransformVX,
                    point.Y
                      - fixture.LabelTransformVY
                      - (
                        (CADObjectCommands.Scale - 0.25)
                        * 12
                        * Math.Pow(0.25 / CADObjectCommands.Scale, 1.5)
                      )
                      - (16 * CADObjectCommands.Scale),
                    0
                  );
                  if (Math.Round(rotation, 1) == 1.6 || Math.Round(rotation, 1) == 4.7)
                  {
                    position = new Point3d(
                      point.X + fixture.LabelTransformHX,
                      point.Y
                        + fixture.LabelTransformHY
                        + (
                          (CADObjectCommands.Scale - 0.25)
                          * 12
                          * Math.Pow(0.25 / CADObjectCommands.Scale, 1.5)
                        ),
                      0
                    );
                    position2 = new Point3d(
                      point.X + fixture.LabelTransformHX,
                      point.Y
                        - fixture.LabelTransformHY
                        - (
                          (CADObjectCommands.Scale - 0.25)
                          * 12
                          * Math.Pow(0.25 / CADObjectCommands.Scale, 1.5)
                        )
                        - (16 * CADObjectCommands.Scale),
                      0
                    );
                  }

                  foreach (ObjectId id in br.AttributeCollection)
                  {
                    AttributeReference attRef = (AttributeReference)
                      tr.GetObject(id, OpenMode.ForWrite);
                    if (attRef.Tag == "LIGHTING_NAME")
                    {
                      attRef.Position = position;
                      attRef.Rotation = attRef.Rotation - rotation;
                      attRef.Height = 0.0938 / CADObjectCommands.Scale * 12;
                      attRef.WidthFactor = 1;
                      attRef.HorizontalMode = TextHorizontalMode.TextLeft;
                      attRef.VerticalMode = TextVerticalMode.TextVerticalMid;
                      attRef.Justify = AttachmentPoint.BaseLeft;
                    }
                    if (attRef.Tag == "LIGHTING_CIRCUIT")
                    {
                      attRef.Position = position2;
                      attRef.Rotation = attRef.Rotation - rotation;
                      attRef.Height = 0.0938 / CADObjectCommands.Scale * 12;
                      attRef.WidthFactor = 1;
                      attRef.HorizontalMode = TextHorizontalMode.TextLeft;
                      attRef.VerticalMode = TextVerticalMode.TextVerticalMid;
                      attRef.Justify = AttachmentPoint.BaseLeft;
                    }
                  }
                  tr.Commit();
                }
                if (i == 0)
                {
                  using (Transaction tr = db.TransactionManager.StartTransaction())
                  {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                    BlockTableRecord block = (BlockTableRecord)
                      tr.GetObject(bt["4CFM"], OpenMode.ForRead);
                    double scaleUp = 12 / CADObjectCommands.Scale;
                    BlockJig blockJig = new BlockJig("block", scaleUp);

                    PromptResult res = blockJig.DragMe(block.ObjectId, out point);

                    if (res.Status == PromptStatus.OK)
                    {
                      BlockTableRecord curSpace = (BlockTableRecord)
                        tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                      BlockReference br = new BlockReference(point, block.ObjectId);
                      br.ScaleFactors = new Scale3d(scaleUp, scaleUp, 1);

                      curSpace.AppendEntity(br);

                      blockId = br.Id;

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
                      AttributeDefinition attrDef = new AttributeDefinition();
                      attrDef.Position = new Point3d(
                        br.Position.X + (0.27 * scaleUp),
                        br.Position.Y + (0.12 * scaleUp),
                        0
                      );

                      var textStyle = (TextStyleTableRecord)
                        tr.GetObject(gmepTextStyleId, OpenMode.ForRead);
                      double widthFactor = 1;
                      if (
                        textStyle.FileName.ToLower().Contains("architxt")
                        || textStyle.FileName.ToLower().Contains("a2")
                      )
                      {
                        widthFactor = 0.85;
                      }

                      attrDef.LockPositionInBlock = false;
                      attrDef.Tag = "tag";
                      attrDef.IsMTextAttributeDefinition = false;
                      attrDef.TextString = fixture.Name;
                      attrDef.Justify = AttachmentPoint.MiddleCenter;
                      attrDef.Visible = true;
                      attrDef.Invisible = false;
                      attrDef.Constant = false;
                      attrDef.Height = 0.0938 * scaleUp;
                      attrDef.WidthFactor = widthFactor;
                      attrDef.TextStyleId = gmepTextStyleId;
                      attrDef.Layer = "0";

                      AttributeReference attrRef = new AttributeReference();
                      attrRef.SetAttributeFromBlock(
                        attrDef,
                        Matrix3d.Displacement(
                          new Vector3d(attrDef.Position.X, attrDef.Position.Y, 0)
                        )
                      );
                      br.AttributeCollection.AppendAttribute(attrRef);

                      AttributeDefinition attrDef2 = new AttributeDefinition();
                      attrDef2.Position = new Point3d(
                        br.Position.X + (0.27 * scaleUp),
                        br.Position.Y - (0.12 * scaleUp),
                        0
                      );

                      attrDef2.LockPositionInBlock = false;
                      attrDef2.Tag = "wattage";
                      attrDef2.IsMTextAttributeDefinition = false;
                      attrDef2.TextString = fixture.Wattage.ToString();
                      attrDef2.Justify = AttachmentPoint.MiddleCenter;
                      attrDef2.Visible = true;
                      attrDef2.Invisible = false;
                      attrDef2.Constant = false;
                      attrDef2.Height = 0.0938 * scaleUp;
                      attrDef2.WidthFactor = widthFactor;
                      attrDef2.TextStyleId = gmepTextStyleId;
                      attrDef2.Layer = "0";

                      AttributeReference attrRef2 = new AttributeReference();
                      attrRef2.SetAttributeFromBlock(
                        attrDef2,
                        Matrix3d.Displacement(
                          new Vector3d(attrDef2.Position.X, attrDef2.Position.Y, 0)
                        )
                      );

                      br.AttributeCollection.AppendAttribute(attrRef2);
                      br.Layer = "E-TXT1";
                      tr.AddNewlyCreatedDBObject(br, true);
                      if (fixture.Qty > 1)
                      {
                        GeneralCommands.CreateAndPositionText(
                          tr,
                          $"TYP. OF {fixture.Qty}",
                          "gmep",
                          0.0938 * scaleUp,
                          1,
                          2,
                          "E-TXT1",
                          new Point3d(
                            br.Position.X + (0.27 * scaleUp),
                            br.Position.Y - (0.32 * scaleUp),
                            0
                          ),
                          TextHorizontalMode.TextLeft,
                          TextVerticalMode.TextBase,
                          AttachmentPoint.MiddleCenter
                        );
                      }
                    }
                    else
                    {
                      return;
                    }

                    tr.Commit();
                  }
                }
              }
              catch (System.Exception ex)
              {
                ed.WriteMessage(ex.ToString());
                Console.WriteLine(ex.ToString());
              }
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
      }
      CreateLightingFixtureListView(true);
    }

    private void RefreshButton_Click(object sender, EventArgs e)
    {
      panelList.Clear();
      lightingFixtureList.Clear();
      lightingControlList.Clear();
      panelList = gmepDb.GetPanels(projectId);
      lightingFixtureList = gmepDb.GetLightingFixtures(projectId);
      lightingControlList = gmepDb.GetLightingControls(projectId);
      CreateLightingFixtureListView(true);
      CreateLightingControlListView(true);
    }
  }
}
