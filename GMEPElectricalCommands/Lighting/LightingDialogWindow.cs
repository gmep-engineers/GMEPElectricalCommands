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
      CreateLightingFixtureListView();

      isLoading = false;
    }

    private Dictionary<string, int> GetNumFixturesOnPlan()
    {
      Dictionary<string, int> fixtureDict = new Dictionary<string, int>();
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;

      Database db = doc.Database;
      Editor ed = doc.Editor;
      Transaction tr = db.TransactionManager.StartTransaction();
      List<PlaceableElectricalEntity> pooledEquipment = new List<PlaceableElectricalEntity>();
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
                if (prop.PropertyName == "gmep_lighting_fixture_id")
                {
                  if (!fixtureDict.ContainsKey(prop.Value as string))
                  {
                    fixtureDict[prop.Value as string] = 1;
                  }
                  else
                  {
                    fixtureDict[prop.Value as string] += 1;
                  }
                }
              }
            }
          }
          catch (Exception e) { }
        }
      }
      return fixtureDict;
    }

    private void CreateLightingFixtureListView(bool updateOnly = false)
    {
      if (updateOnly)
      {
        LightingFixturesListView.Items.Clear();
      }
      LightingFixturesListView.View = View.Details;
      LightingFixturesListView.FullRowSelect = true;
      Dictionary<string, int> fixtureDict = GetNumFixturesOnPlan();
      foreach (LightingFixture fixture in lightingFixtureList)
      {
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
      using (var tr = db.TransactionManager.StartTransaction())
      {
        var btr = (BlockTableRecord)tr.GetObject(spaceId, OpenMode.ForWrite);
        Table tb = new Table();
        tb.TableStyle = db.Tablestyle;
        tb.Position = startPoint;
        int tableRows = lightingFixtureList.Count + 3;
        int tableCols = 11;
        tb.SetSize(tableRows, tableCols);
        tb.SetRowHeight(0.8911);
        tb.Cells[0, 0].TextString = "LIGHTING FIXTURE SCHEDULE";
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
        for (int i = 0; i < lightingFixtureList.Count; i++)
        {
          int row = i + 2;
          tb.Cells[row, 2].TextString = lightingFixtureList[i].Voltage.ToString();
          tb.Cells[row, 3].TextString = lightingFixtureList[i].Qty.ToString();
          tb.Cells[row, 4].TextString = lightingFixtureList[i].Mounting.ToUpper();
          tb.Cells[row, 5].TextString = lightingFixtureList[i].Description.ToUpper();
          tb.Cells[row, 6].TextString = lightingFixtureList[i].Manufacturer.ToUpper();
          tb.Cells[row, 7].TextString = lightingFixtureList[i].ModelNo.ToUpper();
          tb.Cells[row, 8].TextString = "LED";
          tb.Cells[row, 9].TextString = lightingFixtureList[i].Wattage.ToString();
          tb.Cells[row, 10].TextString = lightingFixtureList[i].Notes.ToUpper();
        }
        tb.Cells[tableRows - 1, 0].TextString =
          "NOTES:\n  1) VERIFY WITH OWNER OR ARCHITECT BEFORE PURCHASING THE LIGHTING FIXTURES.\n 2) LIGHTING ABOVE FOOD OR UTENSILS SHALL BE SHATTERPROOF.";
        BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
        btr.AppendEntity(tb);
        tr.AddNewlyCreatedDBObject(tb, true);
        int r = 0;
        foreach (LightingFixture fixture in lightingFixtureList)
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
              attrDef.LockPositionInBlock = false;
              attrDef.Tag = "tag";
              attrDef.IsMTextAttributeDefinition = false;
              attrDef.TextString = fixture.Name;
              attrDef.Justify = AttachmentPoint.MiddleCenter;
              attrDef.Visible = true;
              attrDef.Invisible = false;
              attrDef.Constant = false;
              attrDef.Height = 0.0938;
              attrDef.WidthFactor = 0.85;
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
              attrDef2.WidthFactor = 0.85;
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
    }

    private void CreateLightingFixtureScheduleButton_Click(object sender, EventArgs e)
    {
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
        Point3d startingPoint;
        Document doc = Autodesk
          .AutoCAD
          .ApplicationServices
          .Application
          .DocumentManager
          .MdiActiveDocument;
        Database db = doc.Database;
        Editor ed = doc.Editor;
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          var promptOptions = new PromptPointOptions("\nSelect upper left point:");
          var promptResult = ed.GetPoint(promptOptions);
          if (promptResult.Status == PromptStatus.OK)
            startingPoint = promptResult.Value;
          else
          {
            return;
          }
        }
        CreateLightingFixtureSchedule(doc, db, ed, startingPoint);
      }
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
        Dictionary<string, int> fixtureDict = GetNumFixturesOnPlan();
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

                  BlockTableRecord block = (BlockTableRecord)
                    tr.GetObject(bt[fixture.BlockName], OpenMode.ForRead);
                  BlockJig blockJig = new BlockJig();

                  PromptResult res = blockJig.DragMe(block.ObjectId, out point);

                  if (res.Status == PromptStatus.OK)
                  {
                    BlockTableRecord curSpace = (BlockTableRecord)
                      tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    BlockReference br = new BlockReference(point, block.ObjectId);

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
                      attRef.WidthFactor = 0.85;
                      attRef.HorizontalMode = TextHorizontalMode.TextLeft;
                      attRef.VerticalMode = TextVerticalMode.TextVerticalMid;
                      attRef.Justify = AttachmentPoint.BaseLeft;
                    }
                    if (attRef.Tag == "LIGHTING_CIRCUIT")
                    {
                      attRef.Position = position2;
                      attRef.Rotation = attRef.Rotation - rotation;
                      attRef.Height = 0.0938 / CADObjectCommands.Scale * 12;
                      attRef.WidthFactor = 0.85;
                      attRef.HorizontalMode = TextHorizontalMode.TextLeft;
                      attRef.VerticalMode = TextVerticalMode.TextVerticalMid;
                      attRef.Justify = AttachmentPoint.BaseLeft;
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
      }
      CreateLightingFixtureListView(true);
    }
  }
}
