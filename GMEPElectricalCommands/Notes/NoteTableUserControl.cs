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

    public partial class NoteTableUserControl: UserControl
    {
    public ElectricalKeyedNoteTable KeyedNoteTable { get; set; } = new ElectricalKeyedNoteTable();
    public NoteTableUserControl( ElectricalKeyedNoteTable keyedNoteTable)
      {
         InitializeComponent();
        // Initialize the DataGridView
        KeyedNoteTable = keyedNoteTable;
        this.Load += new EventHandler(NoteTableUserControl_Load);
    }

    private void NoteTableUserControl_Load(object sender, EventArgs e) {
      TableGridView.AutoGenerateColumns = true;
      TableGridView.DataSource = KeyedNoteTable.KeyedNotes;
    }
  }
 
}
