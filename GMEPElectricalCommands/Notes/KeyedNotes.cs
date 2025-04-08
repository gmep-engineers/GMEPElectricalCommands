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
using Autodesk.AutoCAD.Geometry;

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
        TableTabControl.SelectedIndex = TableTabControl.SelectedIndex - 1;
        tabForm.Show();

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
      string sheetName = "Meow";
      using (Transaction tr = db.TransactionManager.StartTransaction()) {
        DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
        foreach (DBDictionaryEntry entry in layoutDict) {
          Layout layout = tr.GetObject(entry.Value, OpenMode.ForRead) as Layout;
          if (layout.Handle.ToString() == sheetId) {
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
    

    private void deleteToolStripMenuItem_Click(object sender, EventArgs e) {
      int tabIndex = (int)deleteToolStripMenuItem.Tag;
      if (tabIndex >= 0 && tabIndex < TableTabControl.TabCount) {
        TabPage tabPage = TableTabControl.TabPages[tabIndex];
        if (tabPage.Controls.Count > 0 && tabPage.Controls[0] is NoteTableUserControl noteTableUserControl) {
          // Access properties or methods of NoteTableUserControl
          ElectricalKeyedNoteTable table = noteTableUserControl.KeyedNoteTable;

          // Remove the table from the dictionary
          if (KeyedNoteTables.ContainsKey(table.SheetId)) {
            KeyedNoteTables[table.SheetId].Remove(table);
            DetermineKeyedNoteIndexes(table.SheetId);
          }
        }
        TableTabControl.TabPages.RemoveAt(tabIndex);

      }
    }
    private void TableTabControl_MouseUp(object sender, MouseEventArgs e) {
      if (e.Button == MouseButtons.Right) {
        for (int i = 0; i < TableTabControl.TabCount; i++) {
          if (TableTabControl.GetTabRect(i).Contains(e.Location)) {
            deleteToolStripMenuItem.Tag = i;
            TabMenu.Show(TableTabControl, e.Location);
            break;
          }
        }
      }
    }
    private void placeToolStripMenuItem_Click(object sender, EventArgs e) {
      int tabIndex = (int)deleteToolStripMenuItem.Tag;
      if (tabIndex >= 0 && tabIndex < TableTabControl.TabCount) {
        TabPage tabPage = TableTabControl.TabPages[tabIndex];
        if (tabPage.Controls.Count > 0 && tabPage.Controls[0] is NoteTableUserControl noteTableUserControl) {
          // Access properties or methods of NoteTableUserControl
          ElectricalKeyedNoteTable table = noteTableUserControl.KeyedNoteTable;
          PlaceTable(table);
        }
      }
    }
    private void PlaceTable(ElectricalKeyedNoteTable noteTable) {
      Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      using (
        DocumentLock docLock =
          Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
      ) {
        Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.WindowState = Autodesk
          .AutoCAD
          .Windows
          .Window
          .State
          .Maximized;
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Window.Focus();

        PromptPointOptions ppo = new PromptPointOptions("\nSpecify insertion point: ");
        PromptPointResult ppr = ed.GetPoint(ppo);
        if (ppr.Status == PromptStatus.OK) {
          Point3d insertionPoint = ppr.Value;
          using (Transaction tr = db.TransactionManager.StartTransaction()) {
            BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
            BlockTableRecord currentSpace = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;


            //Get gmep text style
            TextStyleTable textStyleTable = (TextStyleTable)
                tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead);
            ObjectId sectionTitleStyleId;
            ObjectId gmepStyleId;
            if (textStyleTable.Has("gmep")) {
              gmepStyleId = textStyleTable["gmep"];
            }
            else {
              ed.WriteMessage("\nText style 'gmep' not found. Using default text style.");
              gmepStyleId = doc.Database.Textstyle;
            }
            if (textStyleTable.Has("section title")) {
              sectionTitleStyleId = textStyleTable["section title"];
            }
            else {
              ed.WriteMessage("\nText style 'gmep' not found. Using default text style.");
              sectionTitleStyleId = doc.Database.Textstyle;
            }
            // Create a new table here
            Table table = new Table();
            table.Layer = "E-TEXT"; // Set the layer of the table
            table.SetSize(noteTable.KeyedNotes.Count + 2, 2); // Set the size of the table (rows, columns)

            table.Position = insertionPoint; // Set the position of the table
            table.SetRowHeight(.3); // Set the row height
            table.SetColumnWidth(.3); // Set the column width
            table.Cells[0, 0].TextString = noteTable.Title.ToUpper() + " KEYED NOTES";
            table.Cells[0, 0].TextStyleId = sectionTitleStyleId;
            table.Cells[0, 0].TextHeight = 0.25;
            table.Columns[1].Width = 5;

            for (int i = 0; i < noteTable.KeyedNotes.Count; i++) {
              table.Cells[i + 2, 0].TextString = noteTable.KeyedNotes[i].Index.ToString();
              table.Cells[i + 2, 0].TextStyleId = gmepStyleId;
              table.Cells[i + 2, 1].TextString = noteTable.KeyedNotes[i].Note.ToUpper();
              table.Cells[i + 2, 1].TextStyleId = gmepStyleId;
              table.Cells[i + 2, 1].TextHeight = 0.125;
            }
            currentSpace.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);
            // ...
            tr.Commit();
          }
        }
      }
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
