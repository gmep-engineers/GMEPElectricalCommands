using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Reflection;
using System.IO;
using System.Diagnostics;

namespace ElectricalCommands.T24
{
  public class T24Commands {
    [CommandMethod("LTI")]
    public void LTI() {
      string assemblyLocation = Assembly.GetExecutingAssembly().Location;
      string projectFolder = Path.GetDirectoryName(assemblyLocation);
      string ChromeDriverPath = Path.Combine(projectFolder, "selenium-manager", "windows", "chromedriver.exe");


      ProcessStartInfo startInfo = new ProcessStartInfo {
        WorkingDirectory = Path.GetDirectoryName(ChromeDriverPath)
      };

      Process process
      // Initialize ChromeDriver with the custom service
      using (IWebDriver driver = new ChromeDriver(ChromeDriverPath)) {
        // Your Selenium logic here
      }
    }
  }
}
