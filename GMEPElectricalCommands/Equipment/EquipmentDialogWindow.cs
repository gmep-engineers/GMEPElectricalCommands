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
    }

    private void PopulateList() { }

    private void CreateEquipmentListView()
    {
      equipmentListView.View = View.Details;
      ListViewItem item = new ListViewItem("11", 0);
      item.SubItems.Add("Toaster");
      item.SubItems.Add("General");
      item.SubItems.Add("A");
      item.SubItems.Add("120");
      item.SubItems.Add("1");
      item.SubItems.Add("222,333");
      ListViewItem item2 = new ListViewItem("RTU-72", 0);
      item2.SubItems.Add("A/C unit");
      item2.SubItems.Add("Mechanical");
      item2.SubItems.Add("B");
      item2.SubItems.Add("208");
      item2.SubItems.Add("3");
      item2.SubItems.Add("222,222");
      equipmentListView.Columns.Add("Equip #", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Description", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Category", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Panel", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Voltage", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Phase", -2, HorizontalAlignment.Left);
      equipmentListView.Columns.Add("Location", -2, HorizontalAlignment.Left);

      equipmentListView.Items.Add(item);
      equipmentListView.Items.Add(item2);
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
              foreach (DynamicBlockReferenceProperty prop in pc)
              {
                Console.WriteLine(prop.PropertyName + ": " + prop.Value);
              }
              Console.WriteLine(br.Position);
            }
          }
          catch { }
        }
      }
    }

    private void filterClearButton_Click(object sender, EventArgs e)
    {
      CalculateDistances();
    }
  }
}
