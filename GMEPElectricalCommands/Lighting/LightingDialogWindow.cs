using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using DocumentFormat.OpenXml.Drawing.Charts;
using ElectricalCommands.Equipment;
using GMEPElectricalCommands.GmepDatabase;

namespace ElectricalCommands.Lighting
{
  public partial class LightingDialogWindow : Form
  {
    private List<LightingFixture> lightingFixtureList;
    private List<ListViewItem> lightingFixtureListViewList;
    private string projectId;
    public GmepDatabase gmepDb = new GmepDatabase();
    private List<Equipment.Panel> panelList;
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
      string fileName = Path.GetFileName(doc.Name);
      //string projectNo = Regex.Match(fileName, @"[0-9]{2}-[0-9]{3}").Value;
      string projectNo = "24-123";
      projectId = gmepDb.GetProjectId(projectNo);
      panelList = gmepDb.GetPanels(projectId);
      lightingFixtureList = gmepDb.GetLightingFixtures(projectId);
      CreateLightingFixtureListView();

      isLoading = false;
    }

    private void CreateLightingFixtureListView(bool updateOnly = false)
    {
      if (updateOnly)
      {
        LightingFixturesListView.Clear();
      }
      LightingFixturesListView.View = View.Details;
      LightingFixturesListView.FullRowSelect = true;
      foreach (LightingFixture fixture in lightingFixtureList)
      {
        ListViewItem item = new ListViewItem(fixture.name, 0);
        item.SubItems.Add(fixture.blockName);
        item.SubItems.Add(fixture.voltage.ToString());
        item.SubItems.Add(fixture.qty.ToString());
        item.SubItems.Add(fixture.mounting);
        item.SubItems.Add(fixture.description);
        item.SubItems.Add(fixture.manufacturer);
        item.SubItems.Add(fixture.modelNo);
        item.SubItems.Add("LED");
        item.SubItems.Add(Math.Round(fixture.wattage, 1).ToString());
        item.SubItems.Add(fixture.notes);
        LightingFixturesListView.Items.Add(item);
      }
      if (!updateOnly)
      {
        LightingFixturesListView.Columns.Add("Tag", -2, HorizontalAlignment.Left);
        LightingFixturesListView.Columns.Add("Legend", -2, HorizontalAlignment.Left);
        LightingFixturesListView.Columns.Add("Volt", -2, HorizontalAlignment.Left);
        LightingFixturesListView.Columns.Add("Count", -2, HorizontalAlignment.Left);
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
          tb.Cells[row, 2].TextString = lightingFixtureList[i].voltage.ToString();
          tb.Cells[row, 3].TextString = lightingFixtureList[i].qty.ToString();
          tb.Cells[row, 4].TextString = lightingFixtureList[i].mounting.ToUpper();
          tb.Cells[row, 5].TextString = lightingFixtureList[i].description.ToUpper();
          tb.Cells[row, 6].TextString = lightingFixtureList[i].manufacturer.ToUpper();
          tb.Cells[row, 7].TextString = lightingFixtureList[i].modelNo.ToUpper();
          tb.Cells[row, 8].TextString = "LED";
          tb.Cells[row, 9].TextString = lightingFixtureList[i].wattage.ToString();
          tb.Cells[row, 10].TextString = lightingFixtureList[i].notes.ToUpper();
        }
        tb.Cells[tableRows - 1, 0].TextString =
          "NOTES:\n  1) VERIFY WITH OWNER OR ARCHITECHT BEFORE PURCHASING THE LIGHTING FIXTURES.\n 2) LIGHTING ABOVE FOOD OR UTENSILS SHALL BE SHATTERPROOF.";
        BlockTable bt = (BlockTable)tr.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
        btr.AppendEntity(tb);
        tr.AddNewlyCreatedDBObject(tb, true);
        int r = 0;
        foreach (LightingFixture fixture in lightingFixtureList)
        {
          try
          {
            ObjectId blockId = bt[fixture.blockName];
            if (fixture.emCapable)
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
                acBlkRef.ScaleFactors = new Scale3d(fixture.paperSpaceScale);
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
                acBlkRef.ScaleFactors = new Scale3d(fixture.paperSpaceScale);
                tr.AddNewlyCreatedDBObject(acBlkRef, true);
              }
              ObjectId emBlockId = bt[fixture.blockName + " EM"];
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
                acBlkRef.ScaleFactors = new Scale3d(fixture.paperSpaceScale);
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
                acBlkRef.ScaleFactors = new Scale3d(fixture.paperSpaceScale);
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
              attrDef.TextString = fixture.name;
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
              attrDef2.TextString = fixture.wattage.ToString();
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
  }
}
