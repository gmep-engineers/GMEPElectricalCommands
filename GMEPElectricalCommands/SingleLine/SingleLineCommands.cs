using System.IO;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using ElectricalCommands.SingleLine;

namespace ElectricalCommands.SingleLine
{
  public class SingleLineCommands
  {
    private SingleLineDialogWindow SingleLineWindow;

    [CommandMethod("SINGLELINE")]
    public void SINGLELINE()
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Editor ed = doc.Editor;
      string fileName = Path.GetFileName(doc.Name).Substring(0, 6);
      if (!Regex.IsMatch(fileName, @"[0-9]{2}-[0-9]{3}"))
      {
        ed.WriteMessage("\nFilename invalid format. Filename must begin with GMEP project number.");
        return;
      }
      try
      {
        if (this.SingleLineWindow != null && !this.SingleLineWindow.IsDisposed)
        {
          this.SingleLineWindow.BringToFront();
        }
        else
        {
          this.SingleLineWindow = new SingleLineDialogWindow(this);
          this.SingleLineWindow.InitializeModal();
          this.SingleLineWindow.Show();
        }
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage("Error: " + ex.ToString());
      }
    }
  }
}
