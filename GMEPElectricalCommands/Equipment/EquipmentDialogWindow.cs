using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using GMEPElectricalCommands.GmepDatabase;

namespace ElectricalCommands.Equipment
{
  public struct Equipment
  {
    public string equipId,
      fedFromId,
      objectId;
    public string equipNo,
      description,
      category,
      fedFromName;
    public int voltage;
    public bool is3Phase;
    public int feederDistance;
    public Point3d pos;

    public Equipment(
      string eqId = "",
      string ffId = "",
      string eqName = "",
      string desc = "",
      string cat = "",
      string fdrName = "",
      int volts = 0,
      bool is3Ph = false,
      int fdrDist = 0,
      double xLoc = 0,
      double yLoc = 0
    )
    {
      equipId = eqId;
      fedFromId = ffId;
      equipNo = eqName;
      description = desc;
      category = cat;
      fedFromName = fdrName;
      voltage = volts;
      is3Phase = is3Ph;
      pos = new Point3d(xLoc, yLoc, 0);
      feederDistance = fdrDist;
    }
  }

  public struct Feeder
  {
    public string feederId,
      name,
      fedFromId,
      type;
    public Point3d pos;
    public int feederDistance;

    public Feeder(
      string fdrId,
      string n,
      string ffId,
      string t,
      int fdrDist = 0,
      double xLoc = 0,
      double yLoc = 0
    )
    {
      feederId = fdrId;
      name = n;
      fedFromId = ffId;
      type = t;
      feederDistance = fdrDist;
      pos = new Point3d(xLoc, yLoc, 0);
    }
  }

  public partial class EquipmentDialogWindow : Form
  {
    private string filterPanel;
    private string filterVoltage;
    private string filterPhase;
    private string filterEquipNo;
    private string filterCategory;
    private List<ListViewItem> equipmentListViewList;
    private List<Equipment> equipmentList;
    private List<ListViewItem> feederListViewList;
    private List<Feeder> feederList;
    private string projectId;

    public GmepDatabase db = new GmepDatabase();

    public EquipmentDialogWindow(EquipmentCommands EquipCommands)
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
      projectId = db.GetProjectId(projectNo);
      Console.WriteLine(projectId);
      feederList = db.GetFeeders(projectId);
      //projectId = "811962e4-d572-467c-afd2-13cd182fc5ef";
      equipmentList = db.GetEquipment(projectId);
      //equipmentList = new List<Equipment>();
      feederList = new List<Feeder>();
      CreateEquipmentListView();
      CreateFeederListView();
    }

    private void CreateEquipmentListView()
    {
      equipmentListView.View = View.Details;
      equipmentListView.FullRowSelect = true;
      foreach (Equipment equipment in equipmentList)
      {
        ListViewItem item = new ListViewItem(equipment.equipNo, 0);
        item.SubItems.Add(equipment.description);
        item.SubItems.Add(equipment.category);
        item.SubItems.Add(equipment.fedFromName);
        item.SubItems.Add(equipment.feederDistance.ToString());
        item.SubItems.Add(equipment.voltage.ToString());
        item.SubItems.Add(equipment.is3Phase ? "3" : "1");
        item.SubItems.Add(equipment.pos.ToString());
        equipmentListView.Items.Add(item);
      }
      equipmentListView.Columns.Add("Equip #", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Description", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Category", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Panel", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Panel Distance", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Voltage", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Phase", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Location", -2, HorizontalAlignment.Left);
    }

    private void CreateFeederListView()
    {
      feederListView.View = View.Details;
      feederListView.FullRowSelect = true;
      foreach (Feeder feeder in feederList)
      {
        ListViewItem item = new ListViewItem(feeder.name, 0);
        item.SubItems.Add(feeder.type);
        item.SubItems.Add(feeder.pos.ToString());
        feederListView.Items.Add(item);
      }
      feederListView.Columns.Add("Name", -2, HorizontalAlignment.Left);
      feederListView.Columns.Add("Type", -2, HorizontalAlignment.Left);
      feederListView.Columns.Add("Location", -2, HorizontalAlignment.Left);
    }

    private void FilterPanelComboBox_SelectedIndexChanged(object sender, EventArgs e)
    {
      filterPanel = GeneralCommands.GetComboBoxValue(filterPanelComboBox);
    }

    private void FilterVoltageComboBox_SelectedIndexChanged(object sender, EventArgs e)
    {
      filterVoltage = GeneralCommands.GetComboBoxValue(filterVoltageComboBox);
    }

    private void CalculateDistances()
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;

      Database db = doc.Database;
      Editor ed = doc.Editor;
      Transaction tr = db.TransactionManager.StartTransaction();
      List<Equipment> equipmentList = new List<Equipment>();
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
              bool addEquip = false;
              Equipment eq = new Equipment();
              foreach (DynamicBlockReferenceProperty prop in pc)
              {
                if (prop.PropertyName == "equip_id")
                {
                  addEquip = true;
                  eq.equipId = prop.Value as string;
                }
                if (prop.PropertyName == "fed_from_id")
                {
                  addEquip = true;
                  eq.fedFromId = prop.Value as string;
                }
              }
              eq.pos = br.Position;
              if (addEquip)
              {
                equipmentList.Add(eq);
              }
            }
          }
          catch { }
        }
      }
      for (int i = 0; i < equipmentList.Count; i++)
      {
        for (int j = 0; j < equipmentList.Count; j++)
        {
          if (equipmentList[j].fedFromId == equipmentList[i].equipId)
          {
            Equipment equip = equipmentList[j];
            equip.feederDistance = Convert.ToInt32(
              equipmentList[j].pos.DistanceTo(equipmentList[i].pos)
            );
            Console.WriteLine(equip.feederDistance);
          }
        }
      }
    }

    private void filterClearButton_Click(object sender, EventArgs e)
    {
      CalculateDistances();
    }
  }
}
