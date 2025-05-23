﻿using System;
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
using Dreambuild.AutoCAD;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using System.Buffers;
using DocumentFormat.OpenXml.Vml.Office;

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
      UpdateCAD();
    }
    private void UpdateCAD() {
      Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      using (DocumentLock docLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()) {
        Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.WindowState = Autodesk
          .AutoCAD
          .Windows
          .Window
          .State
          .Maximized;
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Window.Focus();
        ed.WriteMessage("\nUpdating CAD...");

        //Loop through each table in autocad
        using (Transaction tr = db.TransactionManager.StartTransaction()) {
          BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

          //Getting the object id of the attribute definition "A" and the object id of "4KNHEX"
          ObjectId blockId = bt["4KNHEX"];
          BlockTableRecord btr1 = (BlockTableRecord)tr.GetObject(bt["4KNHEX"], OpenMode.ForRead);
          ObjectId attDefId = ObjectId.Null;
          foreach (ObjectId id in btr1) {
            DBObject obj = tr.GetObject(id, OpenMode.ForRead);
            if (obj is AttributeDefinition attDef && attDef.Tag.ToUpper() == "A") {
              attDefId = id;
              break;
            }
          }
          //Getting the text style id of "gmep"
          TextStyleTable textStyleTable = (TextStyleTable)
                tr.GetObject(doc.Database.TextStyleTableId, OpenMode.ForRead);
          ObjectId gmepStyleId;
          if (textStyleTable.Has("gmep")) {
            gmepStyleId = textStyleTable["gmep"];
          }
          else {
            ed.WriteMessage("\nText style 'gmep' not found. Using default text style.");
            gmepStyleId = doc.Database.Textstyle;
          }
         

          //Iterating through each block table record
          foreach (ObjectId btrId in bt) {
            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
            foreach (ObjectId entId in btr) {
              Entity ent = tr.GetObject(entId, OpenMode.ForWrite) as Entity;
              if (ent is Table) {
                Table table = (Table)ent;
                // Check if the table has an extension dictionary
                if (table.ExtensionDictionary == ObjectId.Null) {
                  continue;
                }
                DBDictionary extDict = (DBDictionary)tr.GetObject(table.ExtensionDictionary, OpenMode.ForRead);
                if (extDict != null) {
                  if (extDict.Contains("gmep_keyed_note_table_id")) {
                    ObjectId valueId = extDict.GetAt("gmep_keyed_note_table_id");
                    using (Xrecord xRec = (Xrecord)tr.GetObject(valueId, OpenMode.ForRead)) {
                      if (xRec != null && xRec.Data != null && xRec.Data.AsArray().Count() > 0) {
                        TypedValue tv = xRec.Data.AsArray()[0];
                        if (tv.TypeCode == (int)DxfCode.Text) {
                          string noteTableIdString = tv.Value as string;
                          updateTableOnCAD(table, noteTableIdString, blockId, attDefId, gmepStyleId);
                        }
                      }
                    }
                  }
                }
              }
            }
          }
          tr.Commit();
        }
        //loop through each note in autocad
        using (Transaction tr = db.TransactionManager.StartTransaction()) {
          BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
          BlockTableRecord keyedNoteBlock = (BlockTableRecord)tr.GetObject(bt["KEYED NOTE (GENERAL)"], OpenMode.ForRead);
          //Iterating through each block table record
          if (keyedNoteBlock != null) {
            foreach (ObjectId id in keyedNoteBlock.GetAnonymousBlockIds()) {
              if (id.IsValid) {
                using (BlockTableRecord anonymousBtr = tr.GetObject(id, OpenMode.ForRead) as BlockTableRecord) {
                  if (anonymousBtr != null) {
                    foreach (ObjectId objId in anonymousBtr.GetBlockReferenceIds(true, false)) {
                      if (objId.IsValid) {
                        var entity = tr.GetObject(objId, OpenMode.ForWrite) as BlockReference;
                        string keyedNoteId = "";
                        if (entity != null) {
                          foreach (DynamicBlockReferenceProperty prop in entity.DynamicBlockReferencePropertyCollection) {
                            if (prop.PropertyName == "gmep_keyed_note_id") {
                              keyedNoteId = prop.Value as string;
                              break;
                            }
                          }
                          if (keyedNoteId != "0") {
                            updateNoteOnCAD(entity, keyedNoteId, tr);
                          }
                        }
                      }
                    }
                  }
                }
              }
            }
          }
          tr.Commit();
        }
      }
    }
    private void updateTableOnCAD(Table table, string noteTableId, ObjectId blockId, ObjectId attDefId, ObjectId styleId) {
      //update the values of the table
      foreach (var sheetId in KeyedNoteTables.Keys) {
        foreach (var noteTable in KeyedNoteTables[sheetId]) {
          if (noteTable.Id == noteTableId) {
            table.SetSize(noteTable.KeyedNotes.Count + 2, 2);
            for (int i = 0; i < table.Rows.Count - 2; i++) {
              ElectricalKeyedNote note = noteTable.KeyedNotes[i];
              table.Cells[i + 2, 0].BlockTableRecordId = blockId;
              table.Cells[i + 2, 0].SetBlockAttributeValue(attDefId, note.Index.ToString());
              table.Cells[i + 2, 0].Alignment = CellAlignment.TopCenter;
              table.Cells[i + 2, 0].TextStyleId = styleId;
              table.Cells[i + 2, 0].Contents[0].IsAutoScale = false;

              table.Cells[i + 2, 1].TextString = note.Note.ToUpper();
              table.Cells[i + 2, 1].TextHeight = 0.1;
              table.Cells[i + 2, 1].TextStyleId = styleId;
              table.Cells[i + 2, 1].Alignment = CellAlignment.MiddleLeft;
            }
            return;
          }
        }
      }
      table.Erase();
    }
    private void updateNoteOnCAD(BlockReference noteRef, string keyedNoteId, Transaction tr) {
      foreach (var sheetId in KeyedNoteTables.Keys) {
        foreach (var noteTable in KeyedNoteTables[sheetId]) {
          foreach (var note in noteTable.KeyedNotes) {
            if (note.Id == keyedNoteId) {
              foreach (ObjectId attId in noteRef.AttributeCollection) {
                AttributeReference attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForWrite);
                if (attRef.Tag == "A") {
                  attRef.TextString = note.Index.ToString();
                  return;
                }
                break;
              }
            }
          }
        }
      }
      noteRef.Erase();
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
            if (i != TableTabControl.TabCount - 1) {
              deleteToolStripMenuItem.Tag = i;
              TabMenu.Show(TableTabControl, e.Location);
              break;
            }
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
            table.SetRowHeight(.25); // Set the row height
            table.SetColumnWidth(.5); // Set the column width
            table.Cells[0, 0].TextString = noteTable.Title.ToUpper() + " KEYED NOTES";
            table.Cells[0, 0].TextStyleId = sectionTitleStyleId;
            table.Cells[0, 0].TextHeight = 0.25;
            table.Columns[1].Width = 5.5;

            for (int i = 0; i < noteTable.KeyedNotes.Count; i++) {
              //Grabbing the objectid of the attribute definition "A"
              BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt["4KNHEX"], OpenMode.ForRead);
              ObjectId attDefId = ObjectId.Null;
              foreach (ObjectId id in btr) {
                DBObject obj = tr.GetObject(id, OpenMode.ForRead);
                if (obj is AttributeDefinition attDef && attDef.Tag.ToUpper() == "A") {
                  attDefId= id;
                  break;
                }
              }

              //Note Index with block
              table.Cells[i + 2, 0].BlockTableRecordId = bt["4KNHEX"];
              table.Cells[i + 2, 0].SetBlockAttributeValue(attDefId, noteTable.KeyedNotes[i].Index.ToString());
              table.Cells[i + 2, 0].Alignment = CellAlignment.TopCenter;
              table.Cells[i + 2, 0].TextStyleId = gmepStyleId;
              table.Cells[i + 2, 0].Contents[0].IsAutoScale = false;

              //Note Description
              table.Cells[i + 2, 1].TextString = noteTable.KeyedNotes[i].Note.ToUpper();
              table.Cells[i + 2, 1].TextStyleId = gmepStyleId;
              table.Cells[i + 2, 1].TextHeight = 0.1;
              table.Cells[i + 2, 1].Alignment = CellAlignment.MiddleLeft;
              table.Cells[i + 2, 1].Contents[0].IsAutoScale = false;

            }
            currentSpace.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);
            // ...

            //Adding Unique Table Identifier
            if (table.ExtensionDictionary == ObjectId.Null) {
              table.UpgradeOpen();
              table.CreateExtensionDictionary();
            }
            using (DBDictionary extDict = (DBDictionary)tr.GetObject(table.ExtensionDictionary, OpenMode.ForWrite)) {
              Xrecord xRec = new Xrecord();

              xRec.Data = new ResultBuffer(
                  new TypedValue((int)DxfCode.Text, noteTable.Id)
              );
              if (!extDict.Contains("gmep_keyed_note_table_id")) {

                extDict.SetAt("gmep_keyed_note_table_id", xRec);
                tr.AddNewlyCreatedDBObject(xRec, true);
              }
            }
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
