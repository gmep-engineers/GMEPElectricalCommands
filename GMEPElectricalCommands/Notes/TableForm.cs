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
        public string sheetName { get; set; } = string.Empty;
        public string tableType { get; set; } = string.Empty;
        public string sheetId { get; set; } = string.Empty;
        public Dictionary<string, ObjectId> SheetDictionary { get; private set; } = new Dictionary<string, ObjectId>();
        public TableForm()
         {
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
              SheetDictionary[layout.LayoutName] = entry.Value;
              TableSheet.Items.Add(layout.LayoutName);
            }

            tr.Commit();
          }
        }

        private void SheetName_TextChanged(object sender, EventArgs e) {
            sheetName = SheetName.Text;
        }

        private void TableSheet_SelectedIndexChanged(object sender, EventArgs e) {
            if (TableSheet.SelectedItem != null) {
              sheetId = SheetDictionary[TableSheet.SelectedItem.ToString()].ToString();
            }
        }

        private void TableType_SelectedIndexChanged(object sender, EventArgs e) {
              tableType = TableType.SelectedItem.ToString();
        }
        private void button1_Click(object sender, EventArgs e) {
            //Upload the meow meow!
        }
  }
}
