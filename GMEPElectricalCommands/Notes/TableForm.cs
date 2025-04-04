using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;

namespace ElectricalCommands.Notes
{
    public partial class TableForm: Form
    {
        public string tableName { get; set; } = string.Empty;
        public string tableType { get; set; } = string.Empty;
        public string sheetId { get; set; } = string.Empty;
        public KeyedNotes keynotes { get; set; } = null;
        public TableForm(KeyedNotes keynotes)
        {
          this.keynotes = keynotes;
          InitializeComponent();
          PopulateSheetNamesAndIds();
        }
        private void PopulateSheetNamesAndIds() {
          Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
          Database db = doc.Database;

          using (Transaction tr = db.TransactionManager.StartTransaction()) {
            DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

            foreach (DBDictionaryEntry entry in layoutDict) {
              Layout layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
              TableSheet.Items.Add(new SheetItem(layout.LayoutName, entry.Value));
            }

            tr.Commit();
          }
        }

        private void TableName_TextChanged(object sender, EventArgs e) {
            tableName = TableName.Text;
        }

        private void TableSheet_SelectedIndexChanged(object sender, EventArgs e) {
            if (TableSheet.SelectedItem != null) {
              SheetItem selectedSheet = TableSheet.SelectedItem as SheetItem;
              if (selectedSheet != null) {
                sheetId = selectedSheet.Id.ToString();
              }
            }
        }

        private void TableType_SelectedIndexChanged(object sender, EventArgs e) {
              tableType = TableType.SelectedItem.ToString();
        }
        private void button1_Click(object sender, EventArgs e) {
          keynotes.AddTab(tableName, tableType, sheetId);
          this.Close();
        }
  }
  public class SheetItem {
    public string Name { get; set; }
    public ObjectId Id { get; set; }

    public SheetItem(string name, ObjectId id) {
      Name = name;
      Id = id;
    }

    public override string ToString() {
      return Name;
    }
  }
}
