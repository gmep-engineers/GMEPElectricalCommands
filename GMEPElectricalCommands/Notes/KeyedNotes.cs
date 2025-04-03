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

namespace ElectricalCommands.Notes
{
    public partial class KeyedNotes: Form
    {
    // This is a placeholder for the DataTable that will hold the keyed notes
    public List<ElectricalKeyedNote> KeyedNotesCollection { get; set; } = new List<ElectricalKeyedNote>();
    public KeyedNotes()
    { 
        InitializeComponent();
        this.Load += new EventHandler(KeyedNotes_Load);
    }
    public void KeyedNotes_Load(object sender, EventArgs e) {
      // Initialize the DataTable with some sample data
      KeyedNotesGridView.AutoGenerateColumns = true;
      KeyedNotesCollection.Add(new ElectricalKeyedNote() { Id = "1", TableId = "Table1", DateCreated = DateTime.Now, Note = "Sample Note 1", Index = 1 });
      KeyedNotesCollection.Add(new ElectricalKeyedNote() { Id = "2", TableId = "Table2", DateCreated = DateTime.Now, Note = "Sample Note 2", Index = 2 });
      KeyedNotesCollection.Add(new ElectricalKeyedNote() { Id = "3", TableId = "Table3", DateCreated = DateTime.Now, Note = "Sample Note 3", Index = 3 });
      KeyedNotesGridView.DataSource = KeyedNotesCollection;
    }
  }
  public class ElectricalKeyedNote {
    public string Id { get; set; }
    public string TableId { get; set; }
    public DateTime DateCreated { get; set; } = DateTime.Now;
    public string Note { get; set; }
    public int Index { get; set; }
  }

}
