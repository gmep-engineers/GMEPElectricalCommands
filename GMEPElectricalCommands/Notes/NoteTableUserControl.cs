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
      TableGridView.Columns.Add(CreateTextBoxColumn("Index", "Number"));
      TableGridView.Columns.Add(CreateTextBoxColumn("Note", "Note Text"));
      TableGridView.Columns.Add(CreateTextBoxColumn("TableId", "Table Id"));
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
