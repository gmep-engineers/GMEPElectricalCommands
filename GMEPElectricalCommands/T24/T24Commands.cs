using Autodesk.AutoCAD.Runtime;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using Emgu.CV.Ocl;
using OpenQA.Selenium.Support.UI;
namespace ElectricalCommands.T24 {
  public class T24Commands {
    [CommandMethod("LTI")]
    public void LTI() {
      Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
      string assemblyLocation = Assembly.GetExecutingAssembly().Location;
      string projectFolder = Path.GetDirectoryName(assemblyLocation);
      string ChromePath = Path.Combine(projectFolder, "selenium-manager", "windows", "chromedriver.exe");



      using (IWebDriver driver = new ChromeDriver(ChromePath) { }) {
       
      }
    }
  }
}
