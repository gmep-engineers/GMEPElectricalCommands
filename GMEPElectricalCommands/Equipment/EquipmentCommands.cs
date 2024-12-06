using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace ElectricalCommands.Equipment
{
  public class EquipmentCommands
  {
    private EquipmentDialogWindow EquipWindow;

    [CommandMethod("EQUIPLOC")]
    public void EQUIPLOC()
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Editor ed = doc.Editor;
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
