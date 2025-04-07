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
using GMEPElectricalCommands.GmepDatabase;

namespace ElectricalCommands.Notes
{
  public partial class KeyedNotes: Form
  {
    private string projectId;

    public GmepDatabase gmepDb = new GmepDatabase();
    public Dictionary<string, ObservableCollection<ElectricalKeyedNoteTable>> KeyedNoteTables { get; set; } = new Dictionary<string, ObservableCollection<ElectricalKeyedNoteTable>>();
    public KeyedNotes()
    {
      projectId = gmepDb.GetProjectId(CADObjectCommands.GetProjectNoFromFileName());
      KeyedNoteTables = gmepDb.GetKeyedNoteTables(projectId);
      InitializeComponent();
      this.Load += new EventHandler(TabControl_Load);
    }

    public void TabControl_Load(object sender, EventArgs e) {
      //Add 'New' Tab
      AddNewTabButton();

      //Add All Existing Tabs
      foreach (var sheetId in KeyedNoteTables.Keys) {
        foreach (var table in KeyedNoteTables[sheetId]) {
          string sheetName = GetSheetName(sheetId);
          string title = sheetName + " - " + table.Title;
          TabPage newTabPage = new TabPage(title) {
            BackColor = Color.AliceBlue
          };
          TableTabControl.TabPages.Insert(TableTabControl.TabCount - 1, newTabPage);
          TableTabControl.SelectedTab = newTabPage;
          NoteTableUserControl noteTableUserControl = new NoteTableUserControl(table, this);
          noteTableUserControl.Dock = DockStyle.Fill;
          newTabPage.Controls.Add(noteTableUserControl);
        }
        DetermineKeyedNoteIndexes(sheetId);
      }


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
        TableTabControl.SelectedIndex = TableTabControl.SelectedIndex - 1;
      }
    }
    public void AddTab(string tableName, string sheetId) {
      string sheetName = GetSheetName(sheetId);
      string title = sheetName + " - " + tableName;

      TabPage newTabPage = new TabPage(title) {
        BackColor = Color.AliceBlue
      };
      TableTabControl.TabPages.Insert(TableTabControl.TabCount - 1, newTabPage);
      TableTabControl.SelectedTab = newTabPage;

      ElectricalKeyedNoteTable newTable = new ElectricalKeyedNoteTable {
        Title = tableName,
        SheetId = sheetId
      };
      NoteTableUserControl noteTableUserControl = new NoteTableUserControl(newTable, this);
      noteTableUserControl.Dock = DockStyle.Fill;
      newTabPage.Controls.Add(noteTableUserControl);

      if (!KeyedNoteTables.ContainsKey(sheetId)) {
        KeyedNoteTables[sheetId] = new ObservableCollection<ElectricalKeyedNoteTable>();
      }

      KeyedNoteTables[sheetId].Add(newTable);
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
    public void DetermineKeyedNoteIndexes(string sheetId) {
      var index = 0;
      foreach (var table in KeyedNoteTables[sheetId]) {
        foreach (var note in table.KeyedNotes) {
          note.Index = ++index;
        }
      }
    }

    private void Save_Click(object sender, EventArgs e) {
      gmepDb.UpdateKeyNotesTables(projectId, KeyedNoteTables);
    }
  }

  public class ElectricalKeyedNote {
    public string Id { get; set; } = Guid.NewGuid().ToString();
    public string TableId { get; set; }
    public DateTime DateCreated { get; set; } = DateTime.Now;
    public string Note { get; set; }
    public int Index { get; set; } = 0;
  }
  public class ElectricalKeyedNoteTable {
    public string Id { get; set; } = Guid.NewGuid().ToString();
    public string Title { get; set; }
    public string SheetId { get; set; }
    public BindingList<ElectricalKeyedNote> KeyedNotes { get; set; } = new BindingList<ElectricalKeyedNote>();
  }
}
