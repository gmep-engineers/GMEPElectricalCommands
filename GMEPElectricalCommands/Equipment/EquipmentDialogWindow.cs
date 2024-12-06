using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands.Equipment
{
  public struct Equipment
  {
    public string equip_id,
      fed_from_id,
      object_id;
    public Point3d pos;
    public double feederDistance;
  }

  public partial class EquipmentDialogWindow : Form
  {
    private string filterPanel;
    private string filterVoltage;
    private string filterPhase;
    private string filterEquipNo;
    private string filterCategory;
    private List<ListViewItem> equipmentList;

    public EquipmentDialogWindow(EquipmentCommands EquipCommands)
    {
      InitializeComponent();
    }

    public void InitializeModal()
    {
      PopulateList();
      CreateEquipmentListView();
      CreateFeederListView();
    }

    private void PopulateList() { }

    private void CreateEquipmentListView()
    {
      equipmentListView.View = View.Details;
      ListViewItem item = new ListViewItem("11", 0);
      item.SubItems.Add("Toaster");
      item.SubItems.Add("General");
      item.SubItems.Add("A");
      item.SubItems.Add("111");
      item.SubItems.Add("120");
      item.SubItems.Add("1");
      item.SubItems.Add("222,333");
      ListViewItem item2 = new ListViewItem("RTU-72", 0);
      item2.SubItems.Add("A/C unit");
      item2.SubItems.Add("Mechanical");
      item2.SubItems.Add("B");
      item2.SubItems.Add("111");
      item2.SubItems.Add("208");
      item2.SubItems.Add("3");
      item2.SubItems.Add("222,222");
      equipmentListView.Columns.Add("Equip #", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Description", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Category", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Panel", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Panel Distance", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Voltage", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Phase", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Location", -2, HorizontalAlignment.Left);

      equipmentListView.Items.Add(item);
      equipmentListView.Items.Add(item2);
    }

    private void CreateFeederListView()
    {
      feederListView.View = View.Details;
      ListViewItem item = new ListViewItem("A", 0);
      item.SubItems.Add("Panel");
      item.SubItems.Add("333,222");
      item.SubItems.Add("MSB-1");
      item.SubItems.Add("100");
      ListViewItem item2 = new ListViewItem("MSB-1", 0);
      item2.SubItems.Add("Distribution");
      item2.SubItems.Add("333,444");
      item2.SubItems.Add("MS-1");
      item2.SubItems.Add("15");
      feederListView.Columns.Add("Name", -2, HorizontalAlignment.Left);
      feederListView.Columns.Add("Type", -2, HorizontalAlignment.Left);
      feederListView.Columns.Add("Location", -2, HorizontalAlignment.Left);
      feederListView.Columns.Add("Feeder", -2, HorizontalAlignment.Left);
      feederListView.Columns.Add("Feeder Distance", -2, HorizontalAlignment.Left);

      feederListView.Items.Add(item);
      feederListView.Items.Add(item2);
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
                  eq.equip_id = prop.Value as string;
                }
                if (prop.PropertyName == "fed_from_id")
                {
                  addEquip = true;
                  eq.fed_from_id = prop.Value as string;
                }
                if (prop.PropertyName == "object_id")
                {
                  addEquip = true;
                  eq.object_id = prop.Value as string;
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
          if (equipmentList[j].fed_from_id == equipmentList[i].equip_id)
          {
            Equipment equip = equipmentList[j];
            equip.feederDistance = equipmentList[j].pos.DistanceTo(equipmentList[i].pos);
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
