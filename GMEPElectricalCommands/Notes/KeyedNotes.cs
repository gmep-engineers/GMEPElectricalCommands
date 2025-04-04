using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;

namespace ElectricalCommands.Notes
{
  public partial class KeyedNotes: Form
  {
    public ObservableCollection<ElectricalKeyedNoteTable> KeyedNoteTables { get; set; } = new ObservableCollection<ElectricalKeyedNoteTable>();
    public KeyedNotes()
    { 
        InitializeComponent();
        this.Load += new EventHandler(TabControl_Load);
    }

    public void TabControl_Load(object sender, EventArgs e) {
      //Add All Existing Tabs
      //TableTabControl.TabPages.Add(new TabPage("MEOW"));
      //Add 'New' Tab
      AddNewTabButton();
      // Set the initial selected tab to the last one (the one before "ADD NEW" tab)
      TableTabControl.SelectedIndex = TableTabControl.TabCount - 2;
    }

    private void AddNewTabButton() {
      TabPage addNewTabPage = new TabPage("ADD NEW") {
        BackColor = Color.AliceBlue
      };
      TableTabControl.TabPages.Add(addNewTabPage);
      TableTabControl.SelectedIndexChanged += TableTabControl_SelectedIndexChanged;
    }

    private void TableTabControl_SelectedIndexChanged(object sender, EventArgs e) {
      if (TableTabControl.SelectedTab != null && TableTabControl.SelectedTab.Text == "ADD NEW") {
        TableForm tabForm = new TableForm(this);
        tabForm.Show();
      }
    }
    public void AddTab(string tableName, string tableType, string sheetId) {
      string sheetName = GetSheetName(sheetId);
      string title = sheetName + " - " + tableName;

      TabPage newTabPage = new TabPage(title) {
        BackColor = Color.AliceBlue
      };
      TableTabControl.TabPages.Insert(TableTabControl.TabCount - 1, newTabPage);
      TableTabControl.SelectedTab = newTabPage;

      ElectricalKeyedNoteTable newTable = new ElectricalKeyedNoteTable {
        Id = Guid.NewGuid().ToString(),
        Title = tableName,
        TableType = tableType,
        SheetId = sheetId
      };
      NoteTableUserControl noteTableUserControl = new NoteTableUserControl(newTable);
      noteTableUserControl.Dock = DockStyle.Fill;
      newTabPage.Controls.Add(noteTableUserControl);

      KeyedNoteTables.Add(newTable);
    }
    public string GetSheetName(string sheetId) {
      Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      string sheetName = string.Empty;
      using (Transaction tr = db.TransactionManager.StartTransaction()) {
        DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
        foreach (DBDictionaryEntry entry in layoutDict) {
          Layout layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
          if (layout.ObjectId.ToString() == sheetId) {
            sheetName = layout.LayoutName;
            break;
          }
        }
        tr.Commit();
      }
      return sheetName;
    }
  }

  public class ElectricalKeyedNote {
    public string Id { get; set; }
    public string TableId { get; set; }
    public DateTime DateCreated { get; set; } = DateTime.Now;
    public string Note { get; set; }
    public int Index { get; set; }
  }
  public class ElectricalKeyedNoteTable {
    public string Id { get; set; }
    public string Title { get; set; }
    public string TableType { get; set; }
    public string SheetId { get; set; }
    public List<ElectricalKeyedNote> KeyedNotes { get; set; } = new List<ElectricalKeyedNote>();
  }
}
