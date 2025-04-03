using System.IO;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands.Notes {
  public class NotesCommands {
    private KeyedNotes Notes;

    [CommandMethod("KeyedNotes")]
    public void KeyedNotes() {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Editor ed = doc.Editor;
      string fileName = Path.GetFileName(doc.Name).Substring(0, 6);
      if (!Regex.IsMatch(fileName, @"[0-9]{2}-[0-9]{3}")) {
        ed.WriteMessage("\nFilename invalid format. Filename must begin with GMEP project number.");
        return;
      }
      try {
        if (this.Notes != null && !this.Notes.IsDisposed) {
          this.Notes.BringToFront();
        }
        else {
          this.Notes = new KeyedNotes();
          //this.Notes.InitializeModal();
          this.Notes.Show();
        }
      }
      catch (System.Exception ex) {
        ed.WriteMessage("Error: " + ex.ToString());
      }
    }
  }
}
