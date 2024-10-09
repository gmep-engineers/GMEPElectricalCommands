using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ElectricalCommands
{
  public partial class NoteForm : Form
  {
    private PanelUserControl userInterface;
    private List<string> notesStorage;
    private List<string> defaultNotes;

    public NoteForm(PanelUserControl userInterface)
    {
      InitializeComponent();

      this.userInterface = userInterface;

      // Add event listener for form closing
      this.FormClosing += new FormClosingEventHandler(this.NOTE_FORM_Closing);

      match_notes_textbox_to_notes_storage(userInterface.getNotesStorage());
    }

    private void match_notes_textbox_to_notes_storage(List<string> notesStorage)
    {
      NOTES_TEXTBOX.Text = "";

      foreach (var note in notesStorage)
      {
        NOTES_TEXTBOX.AppendText(note + Environment.NewLine);
      }
    }

    private void NOTE_FORM_Closing(object sender, FormClosingEventArgs e)
    {
      var notesStorageCopy = new List<string>();

      foreach (var line in NOTES_TEXTBOX.Lines)
      {
        if (line != "")
        {
          notesStorageCopy.Add(line);
        }
      }

      // Call the update_notes_storage method on the userInterface
      this.userInterface.UpdateNotesStorage(notesStorageCopy);
    }

    private void ADD_NOTE_BUTTON_Click(object sender, EventArgs e)
    {
      // Check if the selected item in the QUICK_ADD_COMBOBOX is not null
      if (QUICK_ADD_COMBOBOX.SelectedItem != null)
      {
        var selectedItem = QUICK_ADD_COMBOBOX.SelectedItem.ToString();
        // Check if the NOTES_TEXTBOX does not already contain the selected item
        if (!NOTES_TEXTBOX.Text.Contains(selectedItem))
        {
          // Add the selected item on a newline to the NOTES_TEXTBOX
          NOTES_TEXTBOX.AppendText(selectedItem + Environment.NewLine);
        }
      }
    }

    private void DONE_BUTTON_Click(object sender, EventArgs e)
    {
      this.Close();
    }
  }
}