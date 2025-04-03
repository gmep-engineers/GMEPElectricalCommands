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
    public List<ElectricalKeyedNote> KeyedNotesCollection { get; set; } = new List<ElectricalKeyedNote>();
      public NoteTableUserControl()
      {
         InitializeComponent();
         // Initialize the DataGridView
         this.Load += new EventHandler(NoteTableUserControl_Load);
    }

    private void NoteTableUserControl_Load(object sender, EventArgs e) {
      TableGridView.AutoGenerateColumns = true;
      TableGridView.DataSource = KeyedNotesCollection;
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
