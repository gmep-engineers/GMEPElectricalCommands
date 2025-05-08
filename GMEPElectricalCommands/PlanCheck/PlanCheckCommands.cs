using System.IO;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands.PlanCheck
{
  public class PlanCheckCommands
  {
    private PlanCheckDialogWindow PlanCheckDialogWindow;

    [CommandMethod("EPlanCheck")]
    public void PlanCheck()
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
        if (this.PlanCheckDialogWindow != null && !this.PlanCheckDialogWindow.IsDisposed)
        {
          this.PlanCheckDialogWindow.BringToFront();
        }
        else
        {
          this.PlanCheckDialogWindow = new PlanCheckDialogWindow();
          this.PlanCheckDialogWindow.InitializeModal();
          this.PlanCheckDialogWindow.Show();
        }
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage("Error: " + ex.ToString());
      }
    }
  }
}
