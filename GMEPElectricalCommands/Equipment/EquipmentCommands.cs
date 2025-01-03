using System.IO;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands.Equipment
{
  public class EquipmentCommands
  {
    private EquipmentDialogWindow EquipWindow;

    [CommandMethod("EQUIP")]
    public void EQUIP()
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
        if (this.EquipWindow != null && !this.EquipWindow.IsDisposed)
        {
          // Bring myForm to the front
          this.EquipWindow.BringToFront();
        }
        else
        {
          // Create a new MainForm if it's not already open
          this.EquipWindow = new EquipmentDialogWindow(this);
          this.EquipWindow.InitializeModal();
          this.EquipWindow.Show();
        }
      }
      catch (System.Exception ex)
      {
        ed.WriteMessage("Error: " + ex.ToString());
      }
    }
  }
}
