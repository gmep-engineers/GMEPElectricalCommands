using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
  }
}
