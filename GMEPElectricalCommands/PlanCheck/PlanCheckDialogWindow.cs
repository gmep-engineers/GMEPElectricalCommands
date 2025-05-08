using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using GMEPElectricalCommands.GmepDatabase;

namespace ElectricalCommands.PlanCheck
{
  public partial class PlanCheckDialogWindow : Form
  {
    List<ObjectId> BlockList;
    List<PlanCheck> PlanCheckList;
    string ProjectId;
    GmepDatabase GmepDb;

    public PlanCheckDialogWindow()
    {
      InitializeComponent();
    }

    public void InitializeModal()
    {
      GmepDb = new GmepDatabase();
      string projectNo = CADObjectCommands.GetProjectNoFromFileName();
      ProjectId = GmepDb.GetProjectId(projectNo);
      InitializeBlockList();
      InitializePlanCheckList();
    }

    private void PopulateBlockListInSpace(Transaction tr, BlockTableRecord space)
    {
      foreach (ObjectId id in space)
      {
        try
        {
          BlockReference br = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
          if (br.IsDynamicBlock)
          {
            BlockTableRecord btrDynamic = (BlockTableRecord)
              tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead);
            if (btrDynamic.Name.StartsWith("GMEP"))
            {
              BlockList.Add(id);
            }
          }
          else
          {
            BlockTableRecord btrStandard = (BlockTableRecord)
              tr.GetObject(br.BlockTableRecord, OpenMode.ForRead);
            if (
              btrStandard.Name.StartsWith("GMEP") || btrStandard.Name.ToUpper().StartsWith("TBLK")
            )
            {
              BlockList.Add(id);
            }
          }
        }
        catch { }
      }
    }

    private void InitializeBlockList()
    {
      BlockList = new List<ObjectId>();
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
        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
        BlockTableRecord btr = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

        var modelSpace = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
        var paperSpace = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.PaperSpace], OpenMode.ForRead);
        PopulateBlockListInSpace(tr, modelSpace);
        PopulateBlockListInSpace(tr, paperSpace);
      }
    }

    private void AddPlanCheckToListView(PlanCheck planCheck)
    {
      PlanCheckList.Add(planCheck);
      ListViewItem item = new ListViewItem(planCheck.Name, 0);
      item.SubItems.Add(planCheck.Description);
      item.SubItems.Add("Not Run");
      PlanCheckListView.Items.Add(item);
    }

    private void InitializePlanCheckList()
    {
      PlanCheckListView.View = View.Details;
      PlanCheckListView.FullRowSelect = true;
      PlanCheckListView.Columns.Add("Name", -2, HorizontalAlignment.Left);
      PlanCheckListView.Columns.Add("Description", -2, HorizontalAlignment.Left);
      PlanCheckListView.Columns.Add("Status", -2, HorizontalAlignment.Left);
      PlanCheckList = new List<PlanCheck>();

      AddPlanCheckToListView(new CheckCaliforniaStamp(ProjectId));
      AddPlanCheckToListView(new CheckKitchenNotes(ProjectId));
    }

    private void RunAllButton_Click(object sender, EventArgs e)
    {
      GmepDb.OpenConnection();
      foreach (PlanCheck pc in PlanCheckList)
      {
        ListViewItem item = PlanCheckListView.FindItemWithText(pc.Name);
        string result = pc.Check(BlockList, GmepDb.GetConnection());
        item.SubItems[item.SubItems.Count - 1].Text = result;
        if (result.StartsWith("PASSED"))
        {
          item.BackColor = Color.DarkCyan;
          item.ForeColor = Color.White;
        }
        if (result.StartsWith("FAILED"))
        {
          item.BackColor = Color.Crimson;
          item.ForeColor = Color.White;
        }
        if (result.StartsWith("N/A"))
        {
          item.BackColor = Color.LightCyan;
          item.ForeColor = Color.Gray;
        }
      }
      GmepDb.CloseConnection();
    }
  }
}
