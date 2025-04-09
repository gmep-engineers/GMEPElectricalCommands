using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using GMEPElectricalCommands.GmepDatabase;

namespace ElectricalCommands.Notes
{
  public partial class NoteTableUserControl : UserControl {
    public ElectricalKeyedNoteTable KeyedNoteTable { get; set; } = new ElectricalKeyedNoteTable();
    public KeyedNotes KeyedNotes { get; set; } = null;
    public NoteTableUserControl(ElectricalKeyedNoteTable keyedNoteTable, KeyedNotes keyedNotes) {
      InitializeComponent();
      KeyedNoteTable = keyedNoteTable;
      KeyedNotes = keyedNotes;
      KeyedNoteTable.KeyedNotes.ListChanged += KeyedNotes_listChanged;
    }
    private void NoteTableUserControl_Load(object sender, EventArgs e) {
      TableGridView.AutoGenerateColumns = false;
      DataGridViewTextBoxColumn indexColumn = CreateTextBoxColumn("Index", "Number");
      indexColumn.ReadOnly = true;
      indexColumn.FillWeight = 1;

      DataGridViewTextBoxColumn noteColumn = CreateTextBoxColumn("Note", "Note Text");
      noteColumn.FillWeight = 8;
      TableGridView.Columns.Add(indexColumn);
      TableGridView.Columns.Add(noteColumn);
      TableGridView.DataSource = KeyedNoteTable.KeyedNotes;
    }
    private DataGridViewTextBoxColumn CreateTextBoxColumn(string dataPropertyName, string headerText) {
      return new DataGridViewTextBoxColumn {
        DataPropertyName = dataPropertyName,
        HeaderText = headerText,
        AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
      };
    }
    private void KeyedNotes_listChanged(object sender, ListChangedEventArgs e) {
      if (e.ListChangedType == ListChangedType.ItemAdded) {
        KeyedNoteTable.KeyedNotes[e.NewIndex].TableId = KeyedNoteTable.Id;
      }
      if (e.ListChangedType == ListChangedType.ItemDeleted || e.ListChangedType == ListChangedType.ItemAdded) {
        KeyedNotes.DetermineKeyedNoteIndexes(KeyedNoteTable.SheetId);
      }
  
    }
    private void Grid_MouseUp(object sender, DataGridViewCellMouseEventArgs e) {
      Point cellPoint = TableGridView.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true).Location;
      Point newPoint = new Point(cellPoint.X + e.Location.X, cellPoint.Y + e.Location.Y);
      if (e.Button == MouseButtons.Right && TableGridView.SelectedRows.Count > 0) {
          gridMenu.Show(TableGridView, newPoint);
      }
    }

    private void deleteToolStripMenuItem_Click(object sender, EventArgs e) {
        foreach (var row in TableGridView.SelectedRows) {
          int tabIndex = TableGridView.Rows.IndexOf((DataGridViewRow)row);
          KeyedNoteTable.KeyedNotes.RemoveAt(tabIndex);
        }
    }
    
   
    private void placeToolStripMenuItem_Click(object sender, EventArgs e) {
      foreach (var row in TableGridView.SelectedRows) {
        int rowIndex = TableGridView.Rows.IndexOf((DataGridViewRow)row);
        PlaceNote(rowIndex);
      }
      
    }

    private void PlaceNote(int rowIndex) {
      Document doc = Autodesk
       .AutoCAD
       .ApplicationServices
       .Application
       .DocumentManager
       .MdiActiveDocument;
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
        Point3d point = new Point3d();
        ObjectId blockId;
        using (Transaction tr = db.TransactionManager.StartTransaction()) {
          BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
          if (!bt.Has("KEYED NOTE (GENERAL)")) {
            ed.WriteMessage("\nError: Block 'KEYED NOTE (GENERAL)' does not exist in the block table.");
            return;
          }
          BlockTableRecord keyedNoteBlock = (BlockTableRecord)tr.GetObject(bt["KEYED NOTE (GENERAL)"], OpenMode.ForRead);
          BlockJig blockJig = new BlockJig(KeyedNoteTable.KeyedNotes[rowIndex].Index.ToString());
          PromptResult res = blockJig.DragMe(keyedNoteBlock.ObjectId, out point);
          BlockReference br = new BlockReference(point, keyedNoteBlock.ObjectId);
          br.Layer = "E-TEXT";

          if (res.Status == PromptStatus.OK) {
              BlockTableRecord currentSpace = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
              currentSpace.AppendEntity(br);
              tr.AddNewlyCreatedDBObject(br, true);
          }
          else {
            return;
          }
          blockId = br.ObjectId;
          foreach (ObjectId objId in keyedNoteBlock) {
            DBObject obj = tr.GetObject(objId, OpenMode.ForRead);
            AttributeDefinition attDef = obj as AttributeDefinition;
            if (attDef != null && !attDef.Constant) {
              using (AttributeReference attRef = new AttributeReference()) {
                attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                if (attRef.Tag == "A") {
                  attRef.TextString = KeyedNoteTable.KeyedNotes[rowIndex].Index.ToString();
                }
                br.AttributeCollection.AppendAttribute(attRef);
                tr.AddNewlyCreatedDBObject(attRef, true);
              }
            }
          }
          tr.Commit();
        }
        using (Transaction tr = db.TransactionManager.StartTransaction()) {
          BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
          var modelSpace = (BlockTableRecord)
          tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
          BlockReference br = (BlockReference)tr.GetObject(blockId, OpenMode.ForWrite);
          DynamicBlockReferencePropertyCollection pc = br.DynamicBlockReferencePropertyCollection;
          foreach (DynamicBlockReferenceProperty prop in pc) {
            if (prop.PropertyName == "gmep_keyed_note_id" && prop.Value as string == "0") {
              prop.Value = KeyedNoteTable.KeyedNotes[rowIndex].Id.ToString();
            }
            if (prop.PropertyName == "gmep_keyed_note_table_id" && prop.Value as string == "0") {
              prop.Value = KeyedNoteTable.KeyedNotes[rowIndex].TableId.ToString();
            }
          }
          tr.Commit();
        }
      }
    }


  }
}

