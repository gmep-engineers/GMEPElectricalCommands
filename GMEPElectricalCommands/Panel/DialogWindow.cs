using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.GraphicsInterface;
using ElectricalCommands.ElectricalEntity;
using GMEPElectricalCommands.GmepDatabase;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ElectricalCommands
{
  public partial class MainForm : Form
  {
    private readonly PanelCommands myCommandsInstance;
    private NewPanelForm newPanelForm;
    private readonly List<PanelUserControl> userControls;
    private Document acDoc;
    private readonly string acDocPath;
    private readonly string acDocFileName;
    private readonly string acDocName;
    public bool initialized = false;
    public bool readOnly = false;

    public MainForm(PanelCommands myCommands)
    {
      InitializeComponent();
      this.myCommandsInstance = myCommands;
      this.newPanelForm = new NewPanelForm(this);
      this.userControls = new List<PanelUserControl>();
      this.Shown += MAINFORM_SHOWN;
      this.FormClosing += MAINFORM_CLOSING;
      this.KeyPreview = true;
      this.KeyDown += new KeyEventHandler(MAINFORM_KEYDOWN);
      this.Deactivate += MAINFORM_DEACTIVATE;
      this.acDoc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      this.acDoc.BeginDocumentClose -= new DocumentBeginCloseEventHandler(DocBeginDocClose);
      this.acDoc.BeginDocumentClose += new DocumentBeginCloseEventHandler(DocBeginDocClose);
      this.acDocPath = Path.GetDirectoryName(this.acDoc.Name);
      this.acDocFileName = Path.GetFileNameWithoutExtension(acDoc.Name);
      this.acDocName = this.acDoc.Name;
    }

    private void DocBeginDocClose(object sender, DocumentBeginCloseEventArgs e)
    {
      if (!readOnly)
      {
        SavePanelDataToLocalJsonFile();
        this.acDoc.Database.SaveAs(
          acDocName,
          true,
          DwgVersion.Current,
          this.acDoc.Database.SecurityParameters
        );
      }
    }

    public List<PanelUserControl> RetrieveUserControls()
    {
      return this.userControls;
    }

    public UserControl FindUserControl(string panelName)
    {
      foreach (PanelUserControl userControl in userControls)
      {
        string userControlName = userControl.Name.Replace("'", "");
        userControlName = userControlName.Replace(" ", "");
        userControlName = userControlName.Replace("-", "");
        userControlName = userControlName.Replace("PANEL", "");
        userControlName = userControlName.Replace("DISTRIB.", "");

        panelName = panelName.Replace("'", "");
        panelName = panelName.Replace(" ", "");
        panelName = panelName.Replace("-", "");
        panelName = panelName.Replace("PANEL", "");
        panelName = panelName.Replace("DISTRIB.", "");

        if (userControlName.ToLower() == panelName.ToLower())
        {
          return userControl;
        }
      }

      return null;
    }

    public void InitializeModal()
    {
      PANEL_TABS.TabPages.Clear();

      List<Dictionary<string, object>> panelStorage = RetrieveSavedPanelData();

      if (panelStorage.Count == 0)
      {
        return;
      }
      else
      {
        MakeTabsAndPopulate(panelStorage);
        this.initialized = true;
      }
    }

    public void DuplicatePanel()
    {
      // Get the currently selected tab
      TabPage selectedTab = PANEL_TABS.SelectedTab;

      // Check if a tab is selected
      if (selectedTab != null)
      {
        // Get the UserControl associated with the selected tab
        PanelUserControl selectedUserControl = (PanelUserControl)selectedTab.Controls[0];

        // Retrieve the panel data from the selected UserControl
        Dictionary<string, object> selectedPanelData = selectedUserControl.RetrieveDataFromModal();

        // Create a deep copy of the selected panel data using serialization
        string jsonData = JsonConvert.SerializeObject(selectedPanelData);
        Dictionary<string, object> duplicatePanelData = JsonConvert.DeserializeObject<
          Dictionary<string, object>
        >(jsonData);

        // Get the original panel name
        string originalPanelName = selectedPanelData["panel"].ToString();

        // Generate a new panel name with a number appended
        string newPanelName = GetNewPanelName(originalPanelName);

        // Update the panel name in the duplicate panel data
        duplicatePanelData["panel"] = newPanelName;

        List<Dictionary<string, object>> newPanelStorage = new List<Dictionary<string, object>>
        {
          duplicatePanelData,
        };

        MakeTabsAndPopulate(newPanelStorage);
      }
    }

    private object DeepCopy(object value)
    {
      if (value is ICloneable cloneable)
      {
        return cloneable.Clone();
      }
      if (value is string || value.GetType().IsValueType)
      {
        return value;
      }
      if (value is Dictionary<string, object> dict)
      {
        Dictionary<string, object> copy = new Dictionary<string, object>();
        foreach (var kvp in dict)
        {
          copy[kvp.Key] = DeepCopy(kvp.Value);
        }
        return copy;
      }
      if (value is List<object> list)
      {
        List<object> copy = new List<object>();
        foreach (var item in list)
        {
          copy.Add(DeepCopy(item));
        }
        return copy;
      }
      if (value is JToken jToken)
      {
        return jToken.DeepClone();
      }
      throw new InvalidOperationException("Unsupported data type in panel data");
    }

    private string GetNewPanelName(string originalPanelName)
    {
      string newPanelName = originalPanelName;

      // Check if the original panel name ends with a number
      int lastNumber = 0;
      int index = originalPanelName.Length - 1;
      while (index >= 0 && char.IsDigit(originalPanelName[index]))
      {
        lastNumber = lastNumber * 10 + (originalPanelName[index] - '0');
        index--;
      }

      if (lastNumber > 0)
      {
        // Increment the last number
        lastNumber++;
        newPanelName = originalPanelName.Substring(0, index + 1) + lastNumber.ToString();
      }
      else
      {
        // Append "2" to the original panel name
        newPanelName = originalPanelName + "2";
      }

      return newPanelName;
    }

    private void MakeTabsAndPopulate(List<Dictionary<string, object>> panelStorage)
    {
      SetupCellsValuesFromPanelData(panelStorage);
      SetupTagsFromPanelData(panelStorage);

      var sortedPanels = panelStorage.OrderBy(panel => panel["panel"].ToString()).ToList();

      foreach (Dictionary<string, object> panel in sortedPanels)
      {
        string panelName = panel["panel"].ToString();
        PanelUserControl userControl = (PanelUserControl)FindUserControl(panelName);
        if (userControl == null)
        {
          continue;
        }
        userControl.AddListeners();
        userControl.ConfigureDistributionPanel(null, null);
        userControl.LinkSubpanels();
        userControl.UpdatePerCellValueChange();
        userControl.SetWarnings();
      }
      foreach (Dictionary<string, object> panel in sortedPanels)
      {
        // Need to loop through the panels again to retrieve values not yet
        // loaded from the previous iteration to ensure all panels link.
        string panelName = panel["panel"].ToString();
        PanelUserControl userControl = (PanelUserControl)FindUserControl(panelName);
        if (userControl == null)
        {
          continue;
        }
        userControl.LinkSubpanels();
        userControl.UpdatePerCellValueChange();
      }
    }

    private void SetupCellsValuesFromPanelData(List<Dictionary<string, object>> panelStorage)
    {
      var sortedPanels = panelStorage.OrderBy(panel => panel["panel"].ToString()).ToList();

      foreach (Dictionary<string, object> panel in sortedPanels)
      {
        string panelName = panel["panel"].ToString();
        bool is3Ph = panel.ContainsKey("phase_c_left");
        PanelUserControl userControl1 = CreateNewPanelTab(panelName, is3Ph);
        userControl1.ClearModalAndRemoveRows(panel);
        userControl1.PopulateModalWithPanelData(panel);
        if (userControl1.readOnly)
        {
          readOnly = true;
          NEW_PANEL_BUTTON.Enabled = !readOnly;
          SAVE_BUTTON.Enabled = !readOnly;
          LOAD_BUTTON.Enabled = !readOnly;
          DUPLICATE_PANEL_BUTTON.Enabled = !readOnly;
        }
        var notes = JsonConvert.DeserializeObject<List<string>>(panel["notes"].ToString());
        userControl1.UpdateNotesStorage(notes);
      }
    }

    private void SetupTagsFromPanelData(List<Dictionary<string, object>> panelStorage)
    {
      foreach (Dictionary<string, object> panel in panelStorage)
      {
        string panelName = panel["panel"].ToString();
        PanelUserControl userControl1 = (PanelUserControl)FindUserControl(panelName);
        if (userControl1 == null)
        {
          continue;
        }
        DataGridView panelGrid = userControl1.RetrievePanelGrid();
        foreach (DataGridViewRow row in panelGrid.Rows)
        {
          int rowIndex = row.Index;
          var tagNames = new Dictionary<string, int>();
          if (panel.ContainsKey("phase_c_left"))
          {
            tagNames = new Dictionary<string, int>()
            {
              { "phase_a_left_tag", 1 },
              { "phase_b_left_tag", 2 },
              { "phase_c_left_tag", 3 },
              { "phase_a_right_tag", 8 },
              { "phase_b_right_tag", 9 },
              { "phase_c_right_tag", 10 },
              { "description_left_tags", 0 },
              { "description_right_tags", 11 },
            };
          }
          else
          {
            tagNames = new Dictionary<string, int>()
            {
              { "phase_a_left_tag", 1 },
              { "phase_b_left_tag", 2 },
              { "phase_a_right_tag", 7 },
              { "phase_b_right_tag", 8 },
              { "description_left_tags", 0 },
              { "description_right_tags", 9 },
            };
          }

          foreach (var tagName in tagNames)
          {
            if (panel.ContainsKey(tagName.Key))
            {
              SetCellValue(panel, tagName.Key, rowIndex, tagName.Value, row);
            }
          }
        }
        userControl1.UpdateCellBackgroundColor();
        userControl1.CalculateBreakerLoad();
      }
    }

    private void SetCellValue(
      Dictionary<string, object> panel,
      string key,
      int rowIndex,
      int cellIndex,
      DataGridViewRow row
    )
    {
      string tag = panel[key].ToString();
      List<string> tagList = JsonConvert.DeserializeObject<List<string>>(tag);

      if (rowIndex < tagList.Count)
      {
        string tagValue = tagList[rowIndex];
        if (!string.IsNullOrEmpty(tagValue))
        {
          if (key.Contains("phase") && tagValue.Contains("="))
          {
            row.Cells[cellIndex].Value = tagValue;
          }
          else if (key.Contains("description"))
          {
            row.Cells[cellIndex].Tag = tagValue;
          }
          else if (key.Contains("phase") && !tagValue.Contains("="))
          {
            row.Cells[cellIndex].Tag = tagValue;
          }
        }
      }
    }

    internal void DeletePanel(PanelUserControl userControl1)
    {
      DialogResult dialogResult = MessageBox.Show(
        "Are you sure you want to delete this panel?",
        "Delete Panel",
        MessageBoxButtons.YesNo
      );
      if (dialogResult == DialogResult.Yes)
      {
        this.userControls.Remove(userControl1);
        PANEL_TABS.TabPages.Remove(userControl1.Parent as TabPage);
        using (
          DocumentLock docLock =
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
        )
        {
          RemovePanelFromStorage(userControl1);
        }
      }
    }

    private void RemovePanelFromStorage(PanelUserControl userControl1)
    {
      var panelData = RetrieveSavedPanelData();
      var userControlName = userControl1.Name.Replace("\'", "").Replace("`", "");
      Dictionary<string, object> panelToRemove = null;
      foreach (Dictionary<string, object> panel in panelData)
      {
        var panelName = panel["panel"].ToString().Replace("\'", "").Replace("`", "");
        if (panelName == userControlName)
        {
          panelToRemove = panel;
        }
        else
        {
          // check if panel is fed from deleted panel
          PanelUserControl p = (PanelUserControl)FindUserControl(panelName);
          p.RemoveFedFrom(userControlName);
          // check if panel is feeding deleted panel
          p.RemoveSubpanel(userControlName, userControl1.Is3Ph());
        }
      }
      if (panelToRemove != null)
      {
        panelData.Remove(panelToRemove);
      }
      StoreDataInJsonFile(panelData);
    }

    internal bool PanelNameExists(string panelName)
    {
      foreach (TabPage tabPage in PANEL_TABS.TabPages)
      {
        if (tabPage.Text.Split(' ')[1].ToLower() == panelName.ToLower())
        {
          return true;
        }
      }
      return false;
    }

    public void AddUserControlToNewTab(UserControl control, TabPage tabPage)
    {
      // Set the user control location and size if needed
      control.Location = new Point(0, 0); // Top-left corner of the tab page
      control.Dock = DockStyle.Fill; // If you want to dock it to fill the tab

      // Add the user control to the controls of the tab page
      tabPage.Controls.Add(control);
    }

    public PanelUserControl CreateNewPanelTab(string tabName, bool is3Ph)
    {
      // if tabname has "PANEL" in it replace it with "Panel"
      if (tabName.Contains("PANEL") || tabName.Contains("Panel"))
      {
        tabName = tabName.Replace("PANEL", "");
        tabName = tabName.Replace("Panel", "");
      }

      // Create a new TabPage
      TabPage newTabPage = new TabPage(tabName);

      // Add the new TabPage to the TabControl
      PANEL_TABS.TabPages.Add(newTabPage);

      // Optional: Select the newly created tab
      PANEL_TABS.SelectedTab = newTabPage;

      // Create a new UserControl
      PanelUserControl userControl1 = new PanelUserControl(
        this.myCommandsInstance,
        this,
        this.newPanelForm,
        tabName,
        is3Ph
      );

      // Add the UserControl to the list of UserControls
      this.userControls.Add(userControl1);

      // Call the method to add the UserControl to the new tab
      AddUserControlToNewTab(userControl1, newTabPage);
      userControl1.SetLoading(false);

      return userControl1;
    }

    public List<Dictionary<string, object>> RetrieveSavedPanelData()
    {
      List<Dictionary<string, object>> allPanelData = new List<Dictionary<string, object>>();

      string savesDirectory = Path.Combine(acDocPath, "Saves");
      string panelSavesDirectory = Path.Combine(savesDirectory, "Panel");

      // Check if the "Saves/Panel" directory exists
      if (Directory.Exists(panelSavesDirectory))
      {
        // Get all JSON files in the directory
        string[] jsonFiles = Directory.GetFiles(panelSavesDirectory, "*.json");

        // If there are any JSON files, find the most recent one
        if (jsonFiles.Length > 0)
        {
          string mostRecentJsonFile = jsonFiles
            .OrderByDescending(f => File.GetLastWriteTime(f))
            .First();

          // Read the JSON data from the file
          string jsonData = File.ReadAllText(mostRecentJsonFile);

          // Deserialize the JSON data to a list of dictionaries
          allPanelData = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(jsonData);
        }
      }

      return allPanelData;
    }

    public void SavePanelDataToLocalJsonFile()
    {
      List<Dictionary<string, object>> panelStorage = new List<Dictionary<string, object>>();

      if (this.acDoc != null && !readOnly)
      {
        foreach (PanelUserControl userControl in this.userControls)
        {
          panelStorage.Add(userControl.RetrieveDataFromModal());
        }

        StoreDataInJsonFile(panelStorage);
      }
    }

    public void StoreDataInJsonFile(List<Dictionary<string, object>> saveData)
    {
      if (readOnly)
      {
        return;
      }
      string savesDirectory = Path.Combine(acDocPath, "Saves");
      string panelSavesDirectory = Path.Combine(savesDirectory, "Panel");

      // Create the "Saves" directory if it doesn't exist
      if (!Directory.Exists(savesDirectory))
      {
        Directory.CreateDirectory(savesDirectory);
      }

      // Create the "Saves/Panel" directory if it doesn't exist
      if (!Directory.Exists(panelSavesDirectory))
      {
        Directory.CreateDirectory(panelSavesDirectory);
      }

      // Create a JSON file name based on the AutoCAD file name and the current timestamp

      string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
      string jsonFileName = acDocFileName + "_" + timestamp + ".json";
      string jsonFilePath = Path.Combine(panelSavesDirectory, jsonFileName);

      // Serialize all the panel data to JSON
      string jsonData = JsonConvert.SerializeObject(saveData, Formatting.Indented);

      // Write the JSON data to the file
      File.WriteAllText(jsonFilePath, jsonData);
    }

    private void MAINFORM_CLOSING(object sender, FormClosingEventArgs e)
    {
      this.acDoc.BeginDocumentClose -= new DocumentBeginCloseEventHandler(DocBeginDocClose);
      if (!readOnly)
      {
        SavePanelDataToLocalJsonFile();
      }
    }

    public void PANEL_NAME_INPUT_TextChanged(
      object sender,
      EventArgs e,
      string input,
      bool distribSect = false
    )
    {
      int selectedIndex = PANEL_TABS.SelectedIndex;

      if (selectedIndex >= 0)
      {
        if (distribSect)
        {
          PANEL_TABS.TabPages[selectedIndex].Text = "DISTRIB. " + input.ToUpper();
        }
        else
        {
          PANEL_TABS.TabPages[selectedIndex].Text = "PANEL " + input.ToUpper();
        }
      }
    }

    public void PANEL_NAME_INPUT_Leave(
      object sender,
      EventArgs e,
      string input,
      string id,
      string fedFrom
    )
    {
      int selectedIndex = PANEL_TABS.SelectedIndex;
      List<Dictionary<string, object>> panelStorage = RetrieveSavedPanelData();
      string oldPanelName = "";
      foreach (Dictionary<string, object> panel in panelStorage)
      {
        if (!String.IsNullOrEmpty(id) && (panel["id"] as string).ToLower() == id.ToLower())
        {
          oldPanelName = (panel["panel"] as string).Replace("'", "");
        }
      }
      if (!String.IsNullOrEmpty(oldPanelName) && !String.IsNullOrEmpty(fedFrom))
      {
        PanelUserControl fedFromUserControl = (PanelUserControl)FindUserControl(fedFrom);
        fedFromUserControl.UpdateSubpanelName("PANEL " + oldPanelName, "PANEL " + input);
      }
      PanelUserControl userControl = (PanelUserControl)FindUserControl(input);
      userControl.UpdateSubpanelFedFrom();
      SavePanelDataToLocalJsonFile();
    }

    private void MAINFORM_DEACTIVATE(object sender, EventArgs e)
    {
      foreach (PanelUserControl userControl in userControls)
      {
        DataGridView panelGrid = userControl.RetrievePanelGrid();
        panelGrid.ClearSelection();
      }
    }

    private void NEW_PANEL_BUTTON_Click(object sender, EventArgs e)
    {
      this.newPanelForm.ShowDialog();
    }

    private void CREATE_ALL_PANELS_BUTTON_Click(object sender, EventArgs e)
    {
      List<PanelUserControl> userControls = RetrieveUserControls();
      List<Dictionary<string, object>> panels = new List<Dictionary<string, object>>();

      foreach (PanelUserControl userControl in userControls)
      {
        Dictionary<string, object> panelData = userControl.RetrieveDataFromModal();
        panels.Add(panelData);
      }

      using (
        DocumentLock docLock =
          Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
      )
      {
        Close();
        myCommandsInstance.CreatePanels(panels);

        Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.WindowState = Autodesk
          .AutoCAD
          .Windows
          .Window
          .State
          .Maximized;
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Window.Focus();
      }
    }

    private void CreateCurrentPanel_Click(object sender, EventArgs e)
    {
      string selectedTabName = PANEL_TABS.SelectedTab.Text;
      PanelUserControl selectedUserControl = (PanelUserControl)FindUserControl(selectedTabName);
      if (selectedUserControl != null)
      {
        selectedUserControl.CreatePanel();
      }
    }

    private void MAINFORM_SHOWN(object sender, EventArgs e)
    {
      // Check if the userControls list is empty
      if (this.userControls.Count == 0)
      {
        // If empty, show newPanelForm as a modal dialog
        newPanelForm.ShowDialog(); // or use appropriate method to show it as modal
      }
    }

    private void HELP_BUTTON_Click(object sender, EventArgs e)
    {
      HelpForm helpForm = new HelpForm();

      helpForm.ShowDialog();
    }

    private void SAVE_BUTTON_Click(object sender, EventArgs e)
    {
      SavePanelDataToLocalJsonFile();
    }

    private void MAINFORM_KEYDOWN(object sender, KeyEventArgs e)
    {
      if (e.Control && e.KeyCode == Keys.S)
      {
        SavePanelDataToLocalJsonFile();
      }
    }

    private void LOAD_BUTTON_Click(object sender, EventArgs e)
    {
      Close();
      // Prompt the user to select a JSON file
      OpenFileDialog openFileDialog = new OpenFileDialog
      {
        Filter = "JSON files (*.json)|*.json",
        Title = "Select a JSON file",
      };

      if (openFileDialog.ShowDialog() == DialogResult.OK)
      {
        // Read the JSON data from the file
        string jsonData = File.ReadAllText(openFileDialog.FileName);

        // Deserialize the JSON data to a list of dictionaries
        List<Dictionary<string, object>> panelData = JsonConvert.DeserializeObject<
          List<Dictionary<string, object>>
        >(jsonData);

        // Save the panel data
        StoreDataInJsonFile(panelData);
      }
    }

    private void DUPLICATE_PANEL_BUTTON_Click(object sender, EventArgs e)
    {
      DuplicatePanel();
    }

    private void LOAD_CALCULATIONS_BUTTON_Click(object sender, EventArgs e)
    {
      using (DocumentLock docLock = this.acDoc.LockDocument())
      {
        CreateLoadCalculationsTable(this.userControls);
      }
    }

    private void HandleMiniBreaker(
      string[] firstPhaseArr,
      string[] secondPhaseArr,
      string[] descriptionArr,
      string[] breakerArr,
      int startIndex,
      ElectricalEntity.Equipment equip
    )
    {
      string phaseLoad = equip.Va.ToString();
      if (equip.CircuitHalf == 1)
      {
        if (equip.LineVoltage > 120)
        {
          // this-a to next-b
          phaseLoad = Math.Round(Convert.ToDouble(equip.Va) / 2, 0).ToString();
          descriptionArr[startIndex] = equip.Name + " " + equip.Description;
          breakerArr[startIndex] = PanelUserControl.GetBreakerSize(
            equip.Mocp > 0 ? equip.Mocp : equip.Fla * 1.25
          );
          if (String.IsNullOrEmpty(breakerArr[startIndex + 1]))
          {
            breakerArr[startIndex + 1] = "0";
          }
          if (String.IsNullOrEmpty(breakerArr[startIndex + 2]))
          {
            breakerArr[startIndex + 2] = "0";
          }
          breakerArr[startIndex + 3] = "2";
          firstPhaseArr[startIndex] = phaseLoad;
          secondPhaseArr[startIndex + 3] = phaseLoad;
        }
        else
        {
          // this-a
          descriptionArr[startIndex] = equip.Name + " " + equip.Description;
          breakerArr[startIndex] = PanelUserControl.GetBreakerSize(
            equip.Mocp > 0 ? equip.Mocp : equip.Fla * 1.25
          );
          if (String.IsNullOrEmpty(breakerArr[startIndex + 1]))
          {
            breakerArr[startIndex + 1] = "0";
          }
          firstPhaseArr[startIndex] = phaseLoad;
        }
      }
      if (equip.CircuitHalf == 2)
      {
        if (equip.LineVoltage > 120)
        {
          // this-b to next-a
          phaseLoad = Math.Round(Convert.ToDouble(equip.Va) / 2, 0).ToString();
          descriptionArr[startIndex + 1] = equip.Name + " " + equip.Description;
          breakerArr[startIndex + 1] = PanelUserControl.GetBreakerSize(
            equip.Mocp > 1 ? equip.Mocp : equip.Fla * 1.25
          );
          if (String.IsNullOrEmpty(breakerArr[startIndex]))
          {
            breakerArr[startIndex] = "0";
          }
          if (String.IsNullOrEmpty(breakerArr[startIndex + 3]))
          {
            breakerArr[startIndex + 3] = "0";
          }
          breakerArr[startIndex + 2] = "2";
          firstPhaseArr[startIndex + 1] = phaseLoad;
          secondPhaseArr[startIndex + 2] = phaseLoad;
        }
        else
        {
          // this-b
          descriptionArr[startIndex + 1] = equip.Name + " " + equip.Description;
          breakerArr[startIndex + 1] = PanelUserControl.GetBreakerSize(
            equip.Mocp > 1 ? equip.Mocp : equip.Fla * 1.25
          );
          if (String.IsNullOrEmpty(breakerArr[startIndex]))
          {
            breakerArr[startIndex] = "0";
          }
          firstPhaseArr[startIndex + 1] = phaseLoad;
        }
      }
    }

    private void AddLoadToCircuit3P(
      JsonPanel3P jsonPanel,
      string startPhase,
      string side,
      ElectricalEntity.Equipment equip
    )
    {
      string[] circuitArr;
      string[] descriptionArr;
      string[] breakerArr;
      if (side == "left")
      {
        circuitArr = jsonPanel.circuit_left;
        descriptionArr = jsonPanel.description_left;
        breakerArr = jsonPanel.breaker_left;
      }
      else
      {
        circuitArr = jsonPanel.circuit_right;
        descriptionArr = jsonPanel.description_right;
        breakerArr = jsonPanel.breaker_right;
      }
      int startIndex = 0;
      foreach (string circuit in circuitArr)
      {
        if (circuit == equip.Circuit.ToString())
        {
          break;
        }
        else
        {
          startIndex++;
        }
      }
      if (startIndex >= circuitArr.Length)
      {
        MessageBox.Show(
          $"{equip.Name} - {equip.Description} could not be added since it is not within the range of the circuit numbers."
        );
        return;
      }
      string[] firstPhaseArr;
      string[] secondPhaseArr;
      string[] thirdPhaseArr;

      if (startPhase == "a")
      {
        if (side == "left")
        {
          firstPhaseArr = jsonPanel.phase_a_left;
          secondPhaseArr = jsonPanel.phase_b_left;
          thirdPhaseArr = jsonPanel.phase_c_left;
        }
        else
        {
          firstPhaseArr = jsonPanel.phase_a_right;
          secondPhaseArr = jsonPanel.phase_b_right;
          thirdPhaseArr = jsonPanel.phase_c_right;
        }
      }
      else if (startPhase == "b")
      {
        if (side == "left")
        {
          firstPhaseArr = jsonPanel.phase_b_left;
          secondPhaseArr = jsonPanel.phase_c_left;
          thirdPhaseArr = jsonPanel.phase_a_left;
        }
        else
        {
          firstPhaseArr = jsonPanel.phase_b_right;
          secondPhaseArr = jsonPanel.phase_c_right;
          thirdPhaseArr = jsonPanel.phase_a_right;
        }
      }
      else
      {
        if (side == "left")
        {
          firstPhaseArr = jsonPanel.phase_c_left;
          secondPhaseArr = jsonPanel.phase_a_left;
          thirdPhaseArr = jsonPanel.phase_b_left;
        }
        else
        {
          firstPhaseArr = jsonPanel.phase_c_right;
          secondPhaseArr = jsonPanel.phase_a_right;
          thirdPhaseArr = jsonPanel.phase_b_right;
        }
      }
      if (equip.CircuitHalf > 0)
      {
        HandleMiniBreaker(
          firstPhaseArr,
          secondPhaseArr,
          descriptionArr,
          breakerArr,
          startIndex,
          equip
        );
        return;
      }
      string phaseLoad = string.Empty;
      if (equip.Pole == 3)
      {
        phaseLoad = Math.Round(Convert.ToDouble(equip.Va) / 1.732, 0).ToString();
        descriptionArr[startIndex] = equip.Name + " " + equip.Description;
        breakerArr[startIndex] =
          equip.Mocp > 1
            ? equip.Mocp.ToString()
            : PanelUserControl.GetBreakerSize(equip.Fla * 1.25);
        breakerArr[startIndex + 2] = "";
        breakerArr[startIndex + 4] = "3";
        firstPhaseArr[startIndex] = phaseLoad;
        secondPhaseArr[startIndex + 2] = phaseLoad;
        thirdPhaseArr[startIndex + 4] = phaseLoad;
      }
      else if (equip.Pole == 2)
      {
        phaseLoad = Math.Round(Convert.ToDouble(equip.Va) / 2, 0).ToString();
        descriptionArr[startIndex] = equip.Name + " " + equip.Description;
        breakerArr[startIndex] =
          equip.Mocp > 1
            ? equip.Mocp.ToString()
            : PanelUserControl.GetBreakerSize(equip.Fla * 1.25);
        breakerArr[startIndex + 2] = "2";
        firstPhaseArr[startIndex] = phaseLoad;
        secondPhaseArr[startIndex + 2] = phaseLoad;
      }
      else
      {
        phaseLoad = equip.Va.ToString();
        descriptionArr[startIndex] = equip.Name + " " + equip.Description;
        breakerArr[startIndex] =
          equip.Mocp > 1
            ? equip.Mocp.ToString()
            : PanelUserControl.GetBreakerSize(equip.Fla * 1.25);
        firstPhaseArr[startIndex] = phaseLoad;
      }
    }

    private void AddLoadToCircuit2P(
      JsonPanel2P jsonPanel,
      string startPhase,
      string side,
      ElectricalEntity.Equipment equip
    )
    {
      string[] circuitArr;
      string[] descriptionArr;
      string[] breakerArr;
      if (side == "left")
      {
        circuitArr = jsonPanel.circuit_left;
        descriptionArr = jsonPanel.description_left;
        breakerArr = jsonPanel.breaker_left;
      }
      else
      {
        circuitArr = jsonPanel.circuit_right;
        descriptionArr = jsonPanel.description_right;
        breakerArr = jsonPanel.breaker_right;
      }
      int startIndex = 0;
      foreach (string circuit in circuitArr)
      {
        if (circuit == equip.Circuit.ToString())
        {
          break;
        }
        else
        {
          startIndex++;
        }
      }
      if (startIndex >= circuitArr.Length)
      {
        MessageBox.Show(
          $"{equip.Name} - {equip.Description} could not be added since it is not within the range of the circuit numbers."
        );
        return;
      }
      string[] firstPhaseArr;
      string[] secondPhaseArr;

      if (startPhase == "a")
      {
        if (side == "left")
        {
          firstPhaseArr = jsonPanel.phase_a_left;
          secondPhaseArr = jsonPanel.phase_b_left;
        }
        else
        {
          firstPhaseArr = jsonPanel.phase_a_right;
          secondPhaseArr = jsonPanel.phase_b_right;
        }
      }
      else
      {
        if (side == "left")
        {
          firstPhaseArr = jsonPanel.phase_b_left;
          secondPhaseArr = jsonPanel.phase_a_left;
        }
        else
        {
          firstPhaseArr = jsonPanel.phase_b_right;
          secondPhaseArr = jsonPanel.phase_a_right;
        }
      }
      if (equip.CircuitHalf > 0)
      {
        HandleMiniBreaker(
          firstPhaseArr,
          secondPhaseArr,
          descriptionArr,
          breakerArr,
          startIndex,
          equip
        );
        return;
      }
      string phaseLoad = string.Empty;
      if (equip.Pole == 2)
      {
        phaseLoad = Math.Round(Convert.ToDouble(equip.Va) / 2, 0).ToString();
        descriptionArr[startIndex] = equip.Name + " " + equip.Description;
        breakerArr[startIndex] =
          equip.Mocp > 1
            ? equip.Mocp.ToString()
            : PanelUserControl.GetBreakerSize(equip.Fla * 1.25);
        breakerArr[startIndex + 2] = "2";
        firstPhaseArr[startIndex] = phaseLoad;
        secondPhaseArr[startIndex + 2] = phaseLoad;
      }
      else
      {
        phaseLoad = equip.Va.ToString();
        descriptionArr[startIndex] = equip.Name + " " + equip.Description;
        breakerArr[startIndex] =
          equip.Mocp > 1
            ? equip.Mocp.ToString()
            : PanelUserControl.GetBreakerSize(equip.Fla * 1.25);
        firstPhaseArr[startIndex] = phaseLoad;
      }
    }

    private void LoadFromDesignTool_ButtonClick(object sender, EventArgs e)
    {
      // get list of panels
      // get list of equipment
      // generate json based on equipment panel-circuit association + lml/lcl
      GmepDatabase gmepDb = new GmepDatabase();
      string projectId = gmepDb.GetProjectId(CADObjectCommands.GetProjectNoFromFileName());
      List<ElectricalEntity.Panel> panels = gmepDb.GetPanels(projectId);
      List<ElectricalEntity.Equipment> equipment = gmepDb.GetEquipment(projectId);
      List<ElectricalEntity.Transformer> transformers = gmepDb.GetTransformers(projectId);
      List<JsonPanel3P> jsonPanels3P = new List<JsonPanel3P>();
      List<JsonPanel2P> jsonPanels2P = new List<JsonPanel2P>();
      List<string> serializedPanels = new List<string>();
      foreach (ElectricalEntity.Panel panel in panels)
      {
        if (panel.NumBreakers == 0)
        {
          panel.NumBreakers = 42;
        }
        if (panel.Phase == 3)
        {
          List<PanelNote> panelNotes = gmepDb.GetPanelNotes(panel.Id);
          JsonPanel3P jsonPanel = new JsonPanel3P();
          jsonPanel.id = panel.Id;
          jsonPanel.read_only = true;
          jsonPanel.main = panel.MainAmpRating.ToString();
          jsonPanel.panel = "'" + panel.Name + "'";
          jsonPanel.location = panel.DesignLocation;
          jsonPanel.voltage1 = panel.Voltage.Substring(0, 3);
          jsonPanel.voltage2 = panel.Voltage.Substring(4, 3);
          jsonPanel.phase = panel.Phase.ToString();
          jsonPanel.wire = panel.Phase == 3 ? "4" : "3";
          jsonPanel.mounting = panel.IsRecessed ? "RECESSED" : "SURFACE";
          jsonPanel.existing = panel.Status.ToUpper();
          jsonPanel.lcl_override = false;
          jsonPanel.lml_override = false;
          jsonPanel.subtotal_a = "0";
          jsonPanel.subtotal_b = "0";
          jsonPanel.subtotal_c = "0";
          jsonPanel.total_va = "0";
          jsonPanel.lcl = 0;
          jsonPanel.lcl125 = 0;
          jsonPanel.lml = 0;
          jsonPanel.lml125 = 0;
          jsonPanel.kva = "0";
          jsonPanel.feeder_amps = "0";
          jsonPanel.custom_title = "";
          jsonPanel.bus_rating = panel.BusAmpRating.ToString();
          jsonPanel.description_left_highlights = new bool[panel.NumBreakers];
          jsonPanel.description_right_highlights = new bool[panel.NumBreakers];
          jsonPanel.breaker_left_highlights = new bool[panel.NumBreakers];
          jsonPanel.breaker_right_highlights = new bool[panel.NumBreakers];
          jsonPanel.description_left = new string[panel.NumBreakers];
          jsonPanel.description_right = new string[panel.NumBreakers];

          jsonPanel.phase_a_left = new string[panel.NumBreakers];
          jsonPanel.phase_a_right = new string[panel.NumBreakers];
          jsonPanel.phase_b_left = new string[panel.NumBreakers];
          jsonPanel.phase_b_right = new string[panel.NumBreakers];
          jsonPanel.phase_c_left = new string[panel.NumBreakers];
          jsonPanel.phase_c_right = new string[panel.NumBreakers];

          jsonPanel.breaker_left = new string[panel.NumBreakers];
          jsonPanel.breaker_right = new string[panel.NumBreakers];

          jsonPanel.circuit_left = new string[panel.NumBreakers];
          jsonPanel.circuit_right = new string[panel.NumBreakers];

          jsonPanel.phase_a_left_tag = new string[panel.NumBreakers / 2];
          jsonPanel.phase_a_right_tag = new string[panel.NumBreakers / 2];
          jsonPanel.phase_b_left_tag = new string[panel.NumBreakers / 2];
          jsonPanel.phase_b_right_tag = new string[panel.NumBreakers / 2];
          jsonPanel.phase_c_left_tag = new string[panel.NumBreakers / 2];
          jsonPanel.phase_c_right_tag = new string[panel.NumBreakers / 2];

          jsonPanel.description_left_tags = new string[panel.NumBreakers / 2];
          jsonPanel.description_right_tags = new string[panel.NumBreakers / 2];
          int numNotes = 0;
          foreach (PanelNote note in panelNotes)
          {
            if (note.Number > numNotes)
            {
              numNotes = note.Number;
            }
          }
          jsonPanel.notes = new string[numNotes];
          jsonPanels3P.Add(jsonPanel);
          for (int i = 0; i < panel.NumBreakers; i++)
          {
            jsonPanel.description_left_highlights[i] = false;

            jsonPanel.description_right_highlights[i] = false;

            jsonPanel.breaker_left_highlights[i] = false;

            jsonPanel.breaker_right_highlights[i] = false;

            jsonPanel.description_left[i] = "SPACE";

            jsonPanel.description_right[i] = "SPACE";

            jsonPanel.phase_a_left[i] = "0";
            jsonPanel.phase_a_right[i] = "0";

            jsonPanel.phase_b_left[i] = "0";
            jsonPanel.phase_b_right[i] = "0";

            jsonPanel.phase_c_left[i] = "0";
            jsonPanel.phase_c_right[i] = "0";

            jsonPanel.breaker_left[i] = string.Empty;
            jsonPanel.breaker_right[i] = string.Empty;

            if (i % 2 == 0)
            {
              jsonPanel.circuit_left[i] = (i + 1).ToString();
              jsonPanel.circuit_right[i] = (i + 2).ToString();
            }
            else
            {
              jsonPanel.circuit_left[i] = string.Empty;
              jsonPanel.circuit_right[i] = string.Empty;
            }
          }
          int noteIndex = 0;
          for (int i = 0; i < panel.NumBreakers / 2; i++)
          {
            jsonPanel.phase_a_left_tag[i] = string.Empty;
            jsonPanel.phase_a_right_tag[i] = string.Empty;

            jsonPanel.phase_b_left_tag[i] = string.Empty;
            jsonPanel.phase_b_right_tag[i] = string.Empty;

            jsonPanel.phase_c_left_tag[i] = string.Empty;
            jsonPanel.phase_c_right_tag[i] = string.Empty;

            jsonPanel.description_left_tags[i] = string.Empty;
            jsonPanel.description_right_tags[i] = string.Empty;
          }
          for (int i = 0; i < panel.NumBreakers; i++)
          {
            List<PanelNote> leftNotes = panelNotes.FindAll(n =>
              n.CircuitNo % 2 == 1 && n.CircuitNo == i + 1
            );
            foreach (PanelNote leftNote in leftNotes)
            {
              for (int j = 0; j < leftNote.Length; j++)
              {
                jsonPanel.description_left_tags[i / 2 + j] += "|" + leftNote.Note;
                jsonPanel.description_left_tags[i / 2 + j] = jsonPanel
                  .description_left_tags[i / 2 + j]
                  .Trim('|');
              }
              if (!jsonPanel.notes.Contains(leftNote.Note))
              {
                jsonPanel.notes[noteIndex] = leftNote.Note;
                noteIndex++;
              }
            }

            List<PanelNote> rightNotes = panelNotes.FindAll(n =>
              n.CircuitNo % 2 == 0 && n.CircuitNo == i + 2
            );
            foreach (PanelNote rightNote in rightNotes)
            {
              for (int j = 0; j < rightNote.Length; j++)
              {
                jsonPanel.description_right_tags[i / 2 + j] += "|" + rightNote.Note;
                jsonPanel.description_right_tags[i / 2 + j] = jsonPanel
                  .description_right_tags[i / 2 + j]
                  .Trim('|');
              }
              if (!jsonPanel.notes.Contains(rightNote.Note))
              {
                jsonPanel.notes[noteIndex] = rightNote.Note;
                noteIndex++;
              }
            }
          }
          foreach (ElectricalEntity.Equipment equip in equipment)
          {
            if (equip.ParentId == panel.Id)
            {
              if (equip.Circuit % 2 == 0)
              {
                // right 3-phase
                if (equip.Circuit % 3 == 2)
                {
                  // phase a right
                  AddLoadToCircuit3P(jsonPanel, "a", "right", equip);
                }
                else if (equip.Circuit % 3 == 1)
                {
                  // phase b right
                  AddLoadToCircuit3P(jsonPanel, "b", "right", equip);
                }
                else
                {
                  // phase c right
                  AddLoadToCircuit3P(jsonPanel, "c", "right", equip);
                }
              }
              else
              {
                // left 3-phase
                if (equip.Circuit % 3 == 1)
                {
                  // phase a left
                  AddLoadToCircuit3P(jsonPanel, "a", "left", equip);
                }
                else if (equip.Circuit % 3 == 2)
                {
                  // phase c left
                  AddLoadToCircuit3P(jsonPanel, "c", "left", equip);
                }
                else
                {
                  // phase b left
                  AddLoadToCircuit3P(jsonPanel, "b", "left", equip);
                }
              }
            }
          }
          foreach (ElectricalEntity.Panel p in panels)
          {
            if (p.ParentId == panel.Id)
            {
              int voltage = int.Parse(p.Voltage.Substring(4, 3));
              ElectricalEntity.Equipment panelLoad = new ElectricalEntity.Equipment(
                p.Id,
                p.NodeId,
                p.ParentId,
                p.ParentName,
                "PANEL",
                p.Name,
                "",
                voltage,
                p.LoadAmperage,
                p.Phase == 3,
                p.ParentDistance,
                p.Location.X,
                p.Location.Y,
                p.BusAmpRating,
                "",
                Convert.ToInt32(p.Kva * 1000),
                48,
                p.Circuit,
                false,
                false,
                p.Status,
                "",
                0
              );
              if (p.Circuit % 2 == 0)
              {
                // right 3-phase
                if (p.Circuit % 3 == 2)
                {
                  // phase a right
                  AddLoadToCircuit3P(jsonPanel, "a", "right", panelLoad);
                }
                else if (p.Circuit % 3 == 1)
                {
                  // phase b right
                  AddLoadToCircuit3P(jsonPanel, "b", "right", panelLoad);
                }
                else
                {
                  // phase c right
                  AddLoadToCircuit3P(jsonPanel, "c", "right", panelLoad);
                }
              }
              else
              {
                // left 3-phase
                if (p.Circuit % 3 == 1)
                {
                  // phase a left
                  AddLoadToCircuit3P(jsonPanel, "a", "left", panelLoad);
                }
                else if (p.Circuit % 3 == 2)
                {
                  // phase c left
                  AddLoadToCircuit3P(jsonPanel, "c", "left", panelLoad);
                }
                else
                {
                  // phase b left
                  AddLoadToCircuit3P(jsonPanel, "b", "left", panelLoad);
                }
              }
            }
          }
          foreach (ElectricalEntity.Transformer t in transformers)
          {
            if (t.ParentId == panel.Id)
            {
              int va = 0;
              double fla = 0;
              string name = String.Empty;
              bool isPanel = false;
              foreach (ElectricalEntity.Panel p in panels)
              {
                if (p.ParentId == t.Id)
                {
                  va = Convert.ToInt32(p.Kva * 1000);
                  fla = p.LoadAmperage;
                  name = p.Name;
                  isPanel = true;
                }
              }
              if (va == 0)
              {
                foreach (ElectricalEntity.Equipment eq in equipment)
                {
                  if (eq.ParentId == t.Id)
                  {
                    va = Convert.ToInt32(eq.LineVoltage * eq.Fla);
                    fla = eq.LoadAmperage;
                    name = eq.Name;
                  }
                }
              }
              ElectricalEntity.Equipment xfmrLoad = new ElectricalEntity.Equipment(
                t.Id,
                t.NodeId,
                t.ParentId,
                t.ParentName,
                isPanel ? "PANEL " : "XFMR " + t.Name + " for ",
                name,
                "",
                Convert.ToInt32(t.LineVoltage),
                fla,
                t.Phase == 3,
                t.ParentDistance,
                t.Location.X,
                t.Location.Y,
                CADObjectCommands.GetMcaFromFla(t.Kva / t.LineVoltage),
                "",
                va,
                0,
                t.Circuit,
                false,
                false,
                t.Status,
                "",
                0
              );
              if (t.Circuit % 2 == 0)
              {
                // right 3-phase
                if (t.Circuit % 3 == 2)
                {
                  // phase a right
                  AddLoadToCircuit3P(jsonPanel, "a", "right", xfmrLoad);
                }
                else if (t.Circuit % 3 == 1)
                {
                  // phase b right
                  AddLoadToCircuit3P(jsonPanel, "b", "right", xfmrLoad);
                }
                else
                {
                  // phase c right
                  AddLoadToCircuit3P(jsonPanel, "c", "right", xfmrLoad);
                }
              }
              else
              {
                // left 3-phase
                if (t.Circuit % 3 == 1)
                {
                  // phase a left
                  AddLoadToCircuit3P(jsonPanel, "a", "left", xfmrLoad);
                }
                else if (t.Circuit % 3 == 2)
                {
                  // phase c left
                  AddLoadToCircuit3P(jsonPanel, "c", "left", xfmrLoad);
                }
                else
                {
                  // phase b left
                  AddLoadToCircuit3P(jsonPanel, "b", "left", xfmrLoad);
                }
              }
            }
          }

          string serializedPanel = JsonConvert.SerializeObject(jsonPanel, Formatting.Indented);
          serializedPanels.Add(serializedPanel);
        }
        else
        {
          List<PanelNote> panelNotes = gmepDb.GetPanelNotes(panel.Id);
          JsonPanel2P jsonPanel = new JsonPanel2P();
          jsonPanel.id = panel.Id;
          jsonPanel.read_only = true;
          jsonPanel.main = panel.MainAmpRating.ToString();
          jsonPanel.panel = "'" + panel.Name + "'";
          jsonPanel.location = panel.DesignLocation;
          jsonPanel.voltage1 = panel.Voltage.Substring(0, 3);
          jsonPanel.voltage2 = panel.Voltage.Substring(4, 3);
          jsonPanel.phase = panel.Phase.ToString();
          jsonPanel.wire = panel.Phase == 3 ? "4" : "3";
          jsonPanel.mounting = panel.IsRecessed ? "RECESSED" : "SURFACE";
          jsonPanel.existing = panel.Status.ToUpper();
          jsonPanel.lcl_override = false;
          jsonPanel.lml_override = false;
          jsonPanel.subtotal_a = "0";
          jsonPanel.subtotal_b = "0";
          jsonPanel.subtotal_c = "0";
          jsonPanel.total_va = "0";
          jsonPanel.lcl = 0;
          jsonPanel.lcl125 = 0;
          jsonPanel.lml = 0;
          jsonPanel.lml125 = 0;
          jsonPanel.kva = "0";
          jsonPanel.feeder_amps = "0";
          jsonPanel.custom_title = "";
          jsonPanel.bus_rating = panel.BusAmpRating.ToString();
          jsonPanel.description_left_highlights = new bool[panel.NumBreakers];
          jsonPanel.description_right_highlights = new bool[panel.NumBreakers];
          jsonPanel.breaker_left_highlights = new bool[panel.NumBreakers];
          jsonPanel.breaker_right_highlights = new bool[panel.NumBreakers];
          jsonPanel.description_left = new string[panel.NumBreakers];
          jsonPanel.description_right = new string[panel.NumBreakers];

          jsonPanel.phase_a_left = new string[panel.NumBreakers];
          jsonPanel.phase_a_right = new string[panel.NumBreakers];
          jsonPanel.phase_b_left = new string[panel.NumBreakers];
          jsonPanel.phase_b_right = new string[panel.NumBreakers];

          jsonPanel.breaker_left = new string[panel.NumBreakers];
          jsonPanel.breaker_right = new string[panel.NumBreakers];

          jsonPanel.circuit_left = new string[panel.NumBreakers];
          jsonPanel.circuit_right = new string[panel.NumBreakers];

          jsonPanel.phase_a_left_tag = new string[panel.NumBreakers / 2];
          jsonPanel.phase_a_right_tag = new string[panel.NumBreakers / 2];
          jsonPanel.phase_b_left_tag = new string[panel.NumBreakers / 2];
          jsonPanel.phase_b_right_tag = new string[panel.NumBreakers / 2];

          jsonPanel.description_left_tags = new string[panel.NumBreakers / 2];
          jsonPanel.description_right_tags = new string[panel.NumBreakers / 2];
          jsonPanel.notes = new string[panelNotes.Count];
          jsonPanels2P.Add(jsonPanel);
          for (int i = 0; i < panel.NumBreakers; i++)
          {
            jsonPanel.description_left_highlights[i] = false;

            jsonPanel.description_right_highlights[i] = false;

            jsonPanel.breaker_left_highlights[i] = false;

            jsonPanel.breaker_right_highlights[i] = false;

            jsonPanel.description_left[i] = "SPACE";

            jsonPanel.description_right[i] = "SPACE";

            jsonPanel.phase_a_left[i] = "0";
            jsonPanel.phase_a_right[i] = "0";

            jsonPanel.phase_b_left[i] = "0";
            jsonPanel.phase_b_right[i] = "0";

            jsonPanel.breaker_left[i] = string.Empty;
            jsonPanel.breaker_right[i] = string.Empty;

            if (i % 2 == 0)
            {
              jsonPanel.circuit_left[i] = (i + 1).ToString();
              jsonPanel.circuit_right[i] = (i + 2).ToString();
            }
            else
            {
              jsonPanel.circuit_left[i] = string.Empty;
              jsonPanel.circuit_right[i] = string.Empty;
            }
          }
          for (int i = 0; i < panel.NumBreakers / 2; i++)
          {
            jsonPanel.phase_a_left_tag[i] = string.Empty;
            jsonPanel.phase_a_right_tag[i] = string.Empty;

            jsonPanel.phase_b_left_tag[i] = string.Empty;
            jsonPanel.phase_b_right_tag[i] = string.Empty;

            jsonPanel.description_left_tags[i] = string.Empty;
            jsonPanel.description_right_tags[i] = string.Empty;
          }
          int noteIndex = 0;
          for (int i = 0; i < panel.NumBreakers; i++)
          {
            List<PanelNote> leftNotes = panelNotes.FindAll(n =>
              n.CircuitNo % 2 == 1 && n.CircuitNo == i + 1
            );
            foreach (PanelNote leftNote in leftNotes)
            {
              for (int j = 0; j < leftNote.Length; j++)
              {
                jsonPanel.description_left_tags[i / 2 + j] += "|" + leftNote.Note;
                jsonPanel.description_left_tags[i / 2 + j] = jsonPanel
                  .description_left_tags[i / 2 + j]
                  .Trim('|');
              }
              if (!jsonPanel.notes.Contains(leftNote.Note))
              {
                jsonPanel.notes[noteIndex] = leftNote.Note;
                noteIndex++;
              }
            }

            List<PanelNote> rightNotes = panelNotes.FindAll(n =>
              n.CircuitNo % 2 == 0 && n.CircuitNo == i + 2
            );
            foreach (PanelNote rightNote in rightNotes)
            {
              for (int j = 0; j < rightNote.Length; j++)
              {
                jsonPanel.description_right_tags[i / 2 + j] += "|" + rightNote.Note;
                jsonPanel.description_right_tags[i / 2 + j] = jsonPanel
                  .description_right_tags[i / 2 + j]
                  .Trim('|');
              }
              if (!jsonPanel.notes.Contains(rightNote.Note))
              {
                jsonPanel.notes[noteIndex] = rightNote.Note;
                noteIndex++;
              }
            }
          }
          foreach (ElectricalEntity.Equipment equip in equipment)
          {
            if (equip.ParentId == panel.Id)
            {
              if (equip.Circuit % 2 == 0)
              {
                // right 1-phase
                if (equip.Circuit % 4 == 2)
                {
                  // phase a
                  AddLoadToCircuit2P(jsonPanel, "a", "right", equip);
                }
                else
                {
                  // phase b
                  AddLoadToCircuit2P(jsonPanel, "b", "right", equip);
                }
              }
              else
              {
                // left 1-phase
                if ((equip.Circuit + 1) % 4 == 2)
                {
                  // phase a left
                  AddLoadToCircuit2P(jsonPanel, "a", "left", equip);
                }
                else
                {
                  // phase b left
                  AddLoadToCircuit2P(jsonPanel, "b", "left", equip);
                }
              }
            }
          }
          foreach (ElectricalEntity.Panel p in panels)
          {
            if (p.ParentId == panel.Id)
            {
              int voltage = int.Parse(p.Voltage.Substring(4, 3));
              ElectricalEntity.Equipment panelLoad = new ElectricalEntity.Equipment(
                p.Id,
                p.NodeId,
                p.ParentId,
                p.ParentName,
                "PANEL",
                p.Name,
                "",
                voltage,
                p.LoadAmperage,
                p.Phase == 3,
                p.ParentDistance,
                p.Location.X,
                p.Location.Y,
                p.BusAmpRating,
                "",
                Convert.ToInt32(p.Kva * 1000),
                48,
                p.Circuit,
                false,
                false,
                p.Status,
                "",
                0
              );
              if (panelLoad.Circuit % 2 == 0)
              {
                // right 1-phase
                if (panelLoad.Circuit % 4 == 2)
                {
                  // phase a
                  AddLoadToCircuit2P(jsonPanel, "a", "right", panelLoad);
                }
                else
                {
                  // phase b
                  AddLoadToCircuit2P(jsonPanel, "b", "right", panelLoad);
                }
              }
              else
              {
                // left 1-phase
                if ((panelLoad.Circuit + 1) % 4 == 2)
                {
                  // phase a left
                  AddLoadToCircuit2P(jsonPanel, "a", "left", panelLoad);
                }
                else
                {
                  // phase b left
                  AddLoadToCircuit2P(jsonPanel, "b", "left", panelLoad);
                }
              }
            }
          }
          foreach (ElectricalEntity.Transformer t in transformers)
          {
            if (t.ParentId == panel.Id)
            {
              int va = 0;
              double fla = 0;
              string name = String.Empty;
              bool isPanel = false;
              foreach (ElectricalEntity.Panel p in panels)
              {
                if (p.ParentId == t.Id)
                {
                  va = Convert.ToInt32(p.Kva * 1000);
                  fla = p.LoadAmperage;
                  isPanel = true;
                  name = p.Name;
                }
              }
              if (va == 0)
              {
                foreach (ElectricalEntity.Equipment eq in equipment)
                {
                  if (eq.ParentId == t.Id)
                  {
                    va = Convert.ToInt32(eq.LineVoltage * eq.Fla);
                    fla = eq.LoadAmperage;
                    name = eq.Name;
                  }
                }
              }
              ElectricalEntity.Equipment xfmrLoad = new ElectricalEntity.Equipment(
                t.Id,
                t.NodeId,
                t.ParentId,
                t.ParentName,
                isPanel ? "PANEL " : "XFMR " + t.Name + " for ",
                name,
                "",
                Convert.ToInt32(t.LineVoltage),
                fla,
                t.Phase == 3,
                t.ParentDistance,
                t.Location.X,
                t.Location.Y,
                CADObjectCommands.GetMcaFromFla(t.Kva / t.LineVoltage),
                "",
                va,
                0,
                t.Circuit,
                false,
                false,
                t.Status,
                "",
                0
              );
              if (xfmrLoad.Circuit % 2 == 0)
              {
                // right 1-phase
                if (xfmrLoad.Circuit % 4 == 2)
                {
                  // phase a
                  AddLoadToCircuit2P(jsonPanel, "a", "right", xfmrLoad);
                }
                else
                {
                  // phase b
                  AddLoadToCircuit2P(jsonPanel, "b", "right", xfmrLoad);
                }
              }
              else
              {
                // left 1-phase
                if ((xfmrLoad.Circuit + 1) % 4 == 2)
                {
                  // phase a left
                  AddLoadToCircuit2P(jsonPanel, "a", "left", xfmrLoad);
                }
                else
                {
                  // phase b left
                  AddLoadToCircuit2P(jsonPanel, "b", "left", xfmrLoad);
                }
              }
            }
          }
          string serializedPanel = JsonConvert.SerializeObject(jsonPanel, Formatting.Indented);
          serializedPanels.Add(serializedPanel);
        }
      }
      if (serializedPanels.Count == 0)
      {
        return;
      }
      try
      {
        string savesDirectory = Path.Combine(acDocPath, "Saves");
        string panelSavesDirectory = Path.Combine(savesDirectory, "Panel");

        // Create the "Saves" directory if it doesn't exist
        if (!Directory.Exists(savesDirectory))
        {
          Directory.CreateDirectory(savesDirectory);
        }

        // Create the "Saves/Panel" directory if it doesn't exist
        if (!Directory.Exists(panelSavesDirectory))
        {
          Directory.CreateDirectory(panelSavesDirectory);
        }

        // Create a JSON file name based on the AutoCAD file name and the current timestamp

        string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
        string jsonFileName = acDocFileName + "_" + timestamp + ".json";
        string jsonFilePath = Path.Combine(panelSavesDirectory, jsonFileName);
        StreamWriter sw = new StreamWriter(jsonFilePath);
        sw.WriteLine("[");
        for (int i = 0; i < serializedPanels.Count - 1; i++)
        {
          sw.WriteLine(serializedPanels[i]);
          sw.WriteLine(",");
        }
        sw.WriteLine(serializedPanels[serializedPanels.Count - 1]);
        sw.WriteLine("]");
        sw.Close();

        // Read the JSON data from the file
        string jsonData = File.ReadAllText(jsonFilePath);

        // Deserialize the JSON data to a list of dictionaries
        List<Dictionary<string, object>> panelData = JsonConvert.DeserializeObject<
          List<Dictionary<string, object>>
        >(jsonData);

        // Save the panel data
        readOnly = false;
        StoreDataInJsonFile(panelData);
        userControls.Clear();
        InitializeModal();
        readOnly = true;
        NEW_PANEL_BUTTON.Enabled = !readOnly;
        SAVE_BUTTON.Enabled = !readOnly;
        LOAD_BUTTON.Enabled = !readOnly;
        DUPLICATE_PANEL_BUTTON.Enabled = !readOnly;
      }
      catch { }
    }

    public static void CreateLoadCalculationsTable(List<PanelUserControl> userControls)
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      // Collect all subpanel names
      HashSet<string> subpanelNames = new HashSet<string>();
      foreach (var userControl in userControls)
      {
        subpanelNames.UnionWith(userControl.GetSubPanels());
      }

      using (DocumentLock docLock = doc.LockDocument())
      {
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
          try
          {
            BlockTableRecord currentSpace = (BlockTableRecord)
              tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            Table table = new Table();

            // Calculate the number of rows (exclude subpanels)
            int rowCount = userControls.Count(uc => !subpanelNames.Contains(uc.GetPanelName())) + 3; // Header + non-subpanel userControls + Total + "IN CONCLUSION:"

            table.TableStyle = db.Tablestyle;
            table.SetSize(rowCount, 3);

            PromptPointResult pr = ed.GetPoint("\nSpecify insertion point: ");
            if (pr.Status != PromptStatus.OK)
              return;
            table.Position = pr.Value;

            // Set layer to "M-TEXT"
            table.Layer = "E-TEXT";

            // Set column widths
            table.Columns[0].Width = 0.5;
            table.Columns[1].Width = 5.0;
            table.Columns[2].Width = 2.5;

            // Set row heights and text properties
            for (int row = 0; row < rowCount; row++)
            {
              table.Rows[row].Height = 0.75;
              for (int col = 0; col < 3; col++)
              {
                Cell cell = table.Cells[row, col];
                cell.TextHeight = (row == 0) ? 0.25 : 0.1;
                cell.TextStyleId = CreateOrGetTextStyle(db, tr, "Archquick");
                cell.Alignment = CellAlignment.MiddleCenter;
              }
            }

            // Populate the table
            table.Cells[0, 0].TextString = "LOAD CALCULATIONS";
            table.MergeCells(CellRange.Create(table, 0, 0, 0, 2));

            double totalKVA = 0;
            int rowIndex = 1;
            int panelCounter = 1;

            foreach (var userControl in userControls)
            {
              string panelName = userControl.GetPanelName();
              string newOrExisting = userControl.GetNewOrExisting();
              if (!subpanelNames.Contains(panelName))
              {
                double kVA = userControl.GetPanelLoad();
                totalKVA += kVA;

                table.Cells[rowIndex, 0].TextString = $"{panelCounter}.";
                table.Cells[rowIndex, 1].TextString = $"{newOrExisting} PANEL '{panelName}'";
                table.Cells[rowIndex, 2].TextString = $"{kVA:F1} KVA";

                rowIndex++;
                panelCounter++;
              }
            }

            int totalRowIndex = rowCount - 2;
            table.Cells[totalRowIndex, 0].TextString = "TOTAL @ 120/208V 3PH 4W";
            table.MergeCells(CellRange.Create(table, totalRowIndex, 0, totalRowIndex, 1));
            table.Cells[totalRowIndex, 2].TextString = $"{totalKVA:F1} KVA";

            int conclusionRowIndex = rowCount - 1;
            table.Cells[conclusionRowIndex, 0].TextString = "IN CONCLUSION:";
            table.MergeCells(CellRange.Create(table, conclusionRowIndex, 0, conclusionRowIndex, 2));

            currentSpace.AppendEntity(table);
            tr.AddNewlyCreatedDBObject(table, true);
            tr.Commit();

            ed.WriteMessage("\nLoad calculations table created successfully.");
          }
          catch (System.Exception ex)
          {
            ed.WriteMessage($"\nError creating load calculations table: {ex.Message}");
            tr.Abort();
          }
        }
      }
    }

    private static ObjectId CreateOrGetTextStyle(Database db, Transaction tr, string styleName)
    {
      TextStyleTable textStyleTable = (TextStyleTable)
        tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);

      if (!textStyleTable.Has(styleName))
      {
        using (TextStyleTableRecord textStyle = new TextStyleTableRecord())
        {
          textStyle.Name = styleName;
          textStyle.Font = new FontDescriptor(styleName, false, false, 0, 0);

          textStyleTable.UpgradeOpen();
          ObjectId textStyleId = textStyleTable.Add(textStyle);
          tr.AddNewlyCreatedDBObject(textStyle, true);

          return textStyleId;
        }
      }
      else
      {
        return textStyleTable[styleName];
      }
    }

    public void UpdateLclLml()
    {
      Document doc = Autodesk
        .AutoCAD
        .ApplicationServices
        .Application
        .DocumentManager
        .MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      LclLmlManager manager = new LclLmlManager();

      // First pass: Collect initial data
      foreach (PanelUserControl userControl in this.userControls)
      {
        LclLmlObject obj = new LclLmlObject(userControl.Name.Replace("'", ""));
        var LclOverride = (int)userControl.GetLclOverride();
        var LmlOverride = (int)userControl.GetLmlOverride();
        obj.LclOverride = LclOverride != 0;
        obj.LmlOverride = LmlOverride != 0;
        obj.Lcl = (LclOverride != 0) ? LclOverride : userControl.CalculateWattageSum("LCL");
        obj.Lml = (LmlOverride != 0) ? LmlOverride : userControl.StoreItemsAndWattage("LML");
        obj.Subpanels = userControl.GetSubPanels();
        manager.List.Add(obj);
      }

      // Second pass: Calculate final LCL and LML values
      foreach (var panel in manager.List)
      {
        CalculateLcl(panel, manager.List);
        CalculateLml(panel, manager.List);
      }

      // Third pass: Update user controls with calculated values
      foreach (PanelUserControl userControl in this.userControls)
      {
        var panelObj = manager.List.Find(p => p.PanelName == userControl.Name.Replace("'", ""));
        if (panelObj != null)
        {
          userControl.UpdateLclLmlLabels(panelObj.Lcl, panelObj.Lml);
        }
      }
    }

    private void CalculateLcl(LclLmlObject panel, List<LclLmlObject> allPanels)
    {
      if (panel.LclOverride)
        return;

      double totalLcl = panel.Lcl;
      foreach (var subpanelName in panel.Subpanels)
      {
        var subpanel = allPanels.Find(p => p.PanelName == subpanelName);
        if (subpanel != null)
        {
          totalLcl += subpanel.Lcl;
        }
      }
      panel.Lcl = totalLcl;
    }

    private void CalculateLml(LclLmlObject panel, List<LclLmlObject> allPanels)
    {
      if (panel.LmlOverride)
        return;

      double maxLml = panel.Lml;
      maxLml = RecursiveCalculateLml(panel, allPanels, maxLml);
      panel.Lml = maxLml;
    }

    private double RecursiveCalculateLml(
      LclLmlObject panel,
      List<LclLmlObject> allPanels,
      double currentMax
    )
    {
      foreach (var subpanelName in panel.Subpanels)
      {
        var subpanel = allPanels.Find(p => p.PanelName == subpanelName);
        if (subpanel != null && !subpanel.LmlOverride)
        {
          currentMax = Math.Max(currentMax, subpanel.Lml);
          currentMax = RecursiveCalculateLml(subpanel, allPanels, currentMax);
        }
      }
      return currentMax;
    }

    public void RemoveFedFrom(string panelName)
    {
      PanelUserControl panel = (PanelUserControl)FindUserControl(panelName);
      if (panel != null)
      {
        panel.RemoveFedFrom(panelName, false);
      }
    }
  }

  public class LclLmlObject
  {
    public double Lcl { get; set; }
    public double Lml { get; set; }
    public bool LclOverride { get; set; }
    public bool LmlOverride { get; set; }
    public List<string> Subpanels { get; set; }
    public string PanelName { get; }

    public LclLmlObject(string panelName)
    {
      Lcl = 0;
      Lml = 0;
      Subpanels = new List<string>();
      PanelName = panelName;
    }
  }

  public class LclLmlManager
  {
    public List<LclLmlObject> List { get; set; }

    public LclLmlManager()
    {
      List = new List<LclLmlObject>();
    }
  }
}
