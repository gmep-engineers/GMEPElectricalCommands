using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.GraphicsInterface;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ElectricalCommands {

  public partial class MainForm : Form {
    private PanelCommands myCommandsInstance;
    private NewPanelForm newPanelForm;
    private List<PanelUserControl> userControls;
    private Document acDoc;
    public bool initialized = false;

    public MainForm(PanelCommands myCommands) {
      InitializeComponent();
      this.myCommandsInstance = myCommands;
      this.newPanelForm = new NewPanelForm(this);
      this.userControls = new List<PanelUserControl>();
      this.Shown += MAINFORM_SHOWN;
      this.FormClosing += MAINFORM_CLOSING;
      this.KeyPreview = true;
      this.KeyDown += new KeyEventHandler(MAINFORM_KEYDOWN);
      this.Deactivate += MAINFORM_DEACTIVATE;
      this.acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      this.acDoc.BeginDocumentClose -= new DocumentBeginCloseEventHandler(docBeginDocClose);
      this.acDoc.BeginDocumentClose += new DocumentBeginCloseEventHandler(docBeginDocClose);
    }

    private void docBeginDocClose(object sender, DocumentBeginCloseEventArgs e) {
      SavePanelDataToLocalJsonFile();
      string fileName = this.acDoc.Name;
      this.acDoc.Database.SaveAs(
          fileName,
          true,
          DwgVersion.Current,
          this.acDoc.Database.SecurityParameters
      );
    }

    public List<PanelUserControl> retrieve_userControls() {
      return this.userControls;
    }

    public UserControl findUserControl(string panelName) {
      foreach (PanelUserControl userControl in userControls) {
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

        if (userControlName.ToLower() == panelName.ToLower()) {
          return userControl;
        }
      }

      return null;
    }

    public void initialize_modal() {
      PANEL_TABS.TabPages.Clear();

      List<Dictionary<string, object>> panelStorage = retrieve_saved_panel_data();

      if (panelStorage.Count == 0) {
        return;
      }
      else {
        MakeTabsAndPopulate(panelStorage);
        this.initialized = true;
      }
    }

    public void DuplicatePanel() {
      // Get the currently selected tab
      TabPage selectedTab = PANEL_TABS.SelectedTab;

      // Check if a tab is selected
      if (selectedTab != null) {
        // Get the UserControl associated with the selected tab
        PanelUserControl selectedUserControl = (PanelUserControl)selectedTab.Controls[0];

        // Retrieve the panel data from the selected UserControl
        Dictionary<string, object> selectedPanelData = selectedUserControl.retrieve_data_from_modal();

        // Create a deep copy of the selected panel data using serialization
        string jsonData = JsonConvert.SerializeObject(selectedPanelData);
        Dictionary<string, object> duplicatePanelData = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonData);

        // Get the original panel name
        string originalPanelName = selectedPanelData["panel"].ToString();

        // Generate a new panel name with a number appended
        string newPanelName = GetNewPanelName(originalPanelName);

        // Update the panel name in the duplicate panel data
        duplicatePanelData["panel"] = newPanelName;

        List<Dictionary<string, object>> newPanelStorage = new List<Dictionary<string, object>> { duplicatePanelData };

        MakeTabsAndPopulate(newPanelStorage);
      }
    }

    private object DeepCopy(object value) {
      if (value is ICloneable cloneable) {
        return cloneable.Clone();
      }
      if (value is string || value.GetType().IsValueType) {
        return value;
      }
      if (value is Dictionary<string, object> dict) {
        Dictionary<string, object> copy = new Dictionary<string, object>();
        foreach (var kvp in dict) {
          copy[kvp.Key] = DeepCopy(kvp.Value);
        }
        return copy;
      }
      if (value is List<object> list) {
        List<object> copy = new List<object>();
        foreach (var item in list) {
          copy.Add(DeepCopy(item));
        }
        return copy;
      }
      if (value is JToken jToken) {
        return jToken.DeepClone();
      }
      throw new InvalidOperationException("Unsupported data type in panel data");
    }

    private string GetNewPanelName(string originalPanelName) {
      string newPanelName = originalPanelName;

      // Check if the original panel name ends with a number
      int lastNumber = 0;
      int index = originalPanelName.Length - 1;
      while (index >= 0 && char.IsDigit(originalPanelName[index])) {
        lastNumber = lastNumber * 10 + (originalPanelName[index] - '0');
        index--;
      }

      if (lastNumber > 0) {
        // Increment the last number
        lastNumber++;
        newPanelName = originalPanelName.Substring(0, index + 1) + lastNumber.ToString();
      }
      else {
        // Append "2" to the original panel name
        newPanelName = originalPanelName + "2";
      }

      return newPanelName;
    }

    private void MakeTabsAndPopulate(List<Dictionary<string, object>> panelStorage) {
      set_up_cell_values_from_panel_data(panelStorage);
      set_up_tags_from_panel_data(panelStorage);

      var sortedPanels = panelStorage.OrderBy(panel => panel["panel"].ToString()).ToList();

      foreach (Dictionary<string, object> panel in sortedPanels) {
        string panelName = panel["panel"].ToString();
        PanelUserControl userControl = (PanelUserControl)findUserControl(panelName);
        if (userControl == null) {
          continue;
        }
        userControl.AddListeners();
        userControl.configure_distribution_panel(null, null);
        userControl.UpdatePerCellValueChange();
      }
    }

    private void set_up_cell_values_from_panel_data(List<Dictionary<string, object>> panelStorage) {
      var sortedPanels = panelStorage.OrderBy(panel => panel["panel"].ToString()).ToList();

      foreach (Dictionary<string, object> panel in sortedPanels) {
        string panelName = panel["panel"].ToString();
        bool is3PH = panel.ContainsKey("phase_c_left");
        PanelUserControl userControl1 = create_new_panel_tab(panelName, is3PH, true);
        userControl1.clear_modal_and_remove_rows(panel);
        userControl1.populate_modal_with_panel_data(panel);
        var notes = JsonConvert.DeserializeObject<List<string>>(panel["notes"].ToString());
        userControl1.update_notes_storage(notes);
      }
    }

    private void set_up_tags_from_panel_data(List<Dictionary<string, object>> panelStorage) {
      foreach (Dictionary<string, object> panel in panelStorage) {
        string panelName = panel["panel"].ToString();
        PanelUserControl userControl1 = (PanelUserControl)findUserControl(panelName);
        if (userControl1 == null) {
          continue;
        }
        DataGridView panelGrid = userControl1.retrieve_panelGrid();
        foreach (DataGridViewRow row in panelGrid.Rows) {
          int rowIndex = row.Index;
          var tagNames = new Dictionary<string, int>();
          if (panel.ContainsKey("phase_c_left")) {
            tagNames = new Dictionary<string, int>()
                        {
                            { "phase_a_left_tag", 1 },
                            { "phase_b_left_tag", 2 },
                            { "phase_c_left_tag", 3 },
                            { "phase_a_right_tag", 8 },
                            { "phase_b_right_tag", 9 },
                            { "phase_c_right_tag", 10 },
                            { "description_left_tags", 0 },
                            { "description_right_tags", 11 }
                        };
          }
          else {
            tagNames = new Dictionary<string, int>()
                        {
                            { "phase_a_left_tag", 1 },
                            { "phase_b_left_tag", 2 },
                            { "phase_a_right_tag", 7 },
                            { "phase_b_right_tag", 8 },
                            { "description_left_tags", 0 },
                            { "description_right_tags", 9 }
                        };
          }

          foreach (var tagName in tagNames) {
            if (panel.ContainsKey(tagName.Key)) {
              set_cell_value(panel, tagName.Key, rowIndex, tagName.Value, row);
            }
          }
        }
        userControl1.update_cell_background_color();
        userControl1.CalculateBreakerLoad();
      }
    }

    private void set_cell_value(Dictionary<string, object> panel, string key, int rowIndex, int cellIndex, DataGridViewRow row) {
      string tag = panel[key].ToString();
      List<string> tagList = JsonConvert.DeserializeObject<List<string>>(tag);

      if (rowIndex < tagList.Count) {
        string tagValue = tagList[rowIndex];
        if (!string.IsNullOrEmpty(tagValue)) {
          if (key.Contains("phase") && tagValue.Contains("=")) {
            row.Cells[cellIndex].Value = tagValue;
          }
          else if (key.Contains("description")) {
            row.Cells[cellIndex].Tag = tagValue;
          }
          else if (key.Contains("phase") && !tagValue.Contains("=")) {
            row.Cells[cellIndex].Tag = tagValue;
          }
        }
      }
    }

    internal void delete_panel(PanelUserControl userControl1) {
      DialogResult dialogResult = MessageBox.Show(
          "Are you sure you want to delete this panel?",
          "Delete Panel",
          MessageBoxButtons.YesNo
      );
      if (dialogResult == DialogResult.Yes) {
        this.userControls.Remove(userControl1);
        PANEL_TABS.TabPages.Remove(userControl1.Parent as TabPage);
        using (
            DocumentLock docLock =
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
        ) {
          remove_panel_from_storage(userControl1);
        }
      }
    }

    private void remove_panel_from_storage(PanelUserControl userControl1) {
      var panelData = retrieve_saved_panel_data();
      var userControlName = userControl1.Name.Replace("\'", "").Replace("`", "");
      foreach (Dictionary<string, object> panel in panelData) {
        var panelName = panel["panel"].ToString().Replace("\'", "").Replace("`", "");
        if (panelName == userControlName) {
          panelData.Remove(panel);
        }
        else {
          // check if panel is fed from deleted panel
          PanelUserControl p = (PanelUserControl)findUserControl(panelName);
          p.RemoveFedFrom(userControlName);
          // check if panel is feeding deleted panel
          p.RemoveSubpanel(userControlName, userControl1.Is3Ph());
        }
      }

      StoreDataInJsonFile(panelData);
    }

    internal bool panel_name_exists(string panelName) {
      foreach (TabPage tabPage in PANEL_TABS.TabPages) {
        if (tabPage.Text.Split(' ')[1].ToLower() == panelName.ToLower()) {
          return true;
        }
      }
      return false;
    }

    public void add_usercontrol_to_new_tab(UserControl control, TabPage tabPage) {
      // Set the user control location and size if needed
      control.Location = new Point(0, 0); // Top-left corner of the tab page
      control.Dock = DockStyle.Fill; // If you want to dock it to fill the tab

      // Add the user control to the controls of the tab page
      tabPage.Controls.Add(control);
    }

    public static void put_in_json_file(object thing) {
      string json = JsonConvert.SerializeObject(thing, Formatting.Indented);
      string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

      var doc = Autodesk
          .AutoCAD
          .ApplicationServices
          .Application
          .DocumentManager
          .MdiActiveDocument;

      string baseFileName = "Test";

      if (string.IsNullOrEmpty(baseFileName)) {
        baseFileName = "panel_data";
      }
      string extension = ".json";
      string path = Path.Combine(desktopPath, baseFileName + extension);

      int count = 1;
      while (File.Exists(path)) {
        string tempFileName = string.Format("{0}({1})", baseFileName, count++);
        path = Path.Combine(desktopPath, tempFileName + extension);
      }

      File.WriteAllText(path, json);
    }

    public PanelUserControl create_new_panel_tab(string tabName, bool is3PH, bool isLoadingData = false) {
      // if tabname has "PANEL" in it replace it with "Panel"
      if (tabName.Contains("PANEL") || tabName.Contains("Panel")) {
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
      PanelUserControl userControl1 = new PanelUserControl(this.myCommandsInstance, this, this.newPanelForm, tabName, is3PH, isLoadingData);

      // Add the UserControl to the list of UserControls
      this.userControls.Add(userControl1);

      // Call the method to add the UserControl to the new tab
      add_usercontrol_to_new_tab(userControl1, newTabPage);

      return userControl1;
    }

    public List<Dictionary<string, object>> retrieve_saved_panel_data() {
      List<Dictionary<string, object>> allPanelData = new List<Dictionary<string, object>>();

      string acDocPath = Path.GetDirectoryName(this.acDoc.Name);
      string savesDirectory = Path.Combine(acDocPath, "Saves");
      string panelSavesDirectory = Path.Combine(savesDirectory, "Panel");

      // Check if the "Saves/Panel" directory exists
      if (Directory.Exists(panelSavesDirectory)) {
        // Get all JSON files in the directory
        string[] jsonFiles = Directory.GetFiles(panelSavesDirectory, "*.json");

        // If there are any JSON files, find the most recent one
        if (jsonFiles.Length > 0) {
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

    public void SavePanelDataToLocalJsonFile() {
      List<Dictionary<string, object>> panelStorage = new List<Dictionary<string, object>>();

      if (this.acDoc != null) {
        using (DocumentLock docLock = this.acDoc.LockDocument()) {
          foreach (PanelUserControl userControl in this.userControls) {
            panelStorage.Add(userControl.retrieve_data_from_modal());
          }

          StoreDataInJsonFile(panelStorage);
        }
      }
    }

    public void StoreDataInJsonFile(List<Dictionary<string, object>> saveData) {
      string acDocPath = Path.GetDirectoryName(this.acDoc.Name);
      string savesDirectory = Path.Combine(acDocPath, "Saves");
      string panelSavesDirectory = Path.Combine(savesDirectory, "Panel");

      // Create the "Saves" directory if it doesn't exist
      if (!Directory.Exists(savesDirectory)) {
        Directory.CreateDirectory(savesDirectory);
      }

      // Create the "Saves/Panel" directory if it doesn't exist
      if (!Directory.Exists(panelSavesDirectory)) {
        Directory.CreateDirectory(panelSavesDirectory);
      }

      // Create a JSON file name based on the AutoCAD file name and the current timestamp
      string acDocFileName = Path.GetFileNameWithoutExtension(acDoc.Name);
      string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
      string jsonFileName = acDocFileName + "_" + timestamp + ".json";
      string jsonFilePath = Path.Combine(panelSavesDirectory, jsonFileName);

      // Serialize all the panel data to JSON
      string jsonData = JsonConvert.SerializeObject(saveData, Formatting.Indented);

      // Write the JSON data to the file
      File.WriteAllText(jsonFilePath, jsonData);
    }

    private void MAINFORM_CLOSING(object sender, FormClosingEventArgs e) {
      this.acDoc.BeginDocumentClose -= new DocumentBeginCloseEventHandler(docBeginDocClose);
      SavePanelDataToLocalJsonFile();
    }

    public void PANEL_NAME_INPUT_TextChanged(object sender, EventArgs e, string input, bool distribSect = false) {
      int selectedIndex = PANEL_TABS.SelectedIndex;

      if (selectedIndex >= 0) {
        if (distribSect) {
          PANEL_TABS.TabPages[selectedIndex].Text = "DISTRIB. " + input.ToUpper();
        }
        else {
          PANEL_TABS.TabPages[selectedIndex].Text = "PANEL " + input.ToUpper();
        }
      }
    }

    public void PANEL_NAME_INPUT_Leave(object sender, EventArgs e, string input, string id, string fedFrom) {
      int selectedIndex = PANEL_TABS.SelectedIndex;
      List<Dictionary<string, object>> panelStorage = retrieve_saved_panel_data();
      string oldPanelName = "";
      foreach (Dictionary<string, object> panel in panelStorage) {
        if (!String.IsNullOrEmpty(id) && (panel["id"] as string).ToLower() == id.ToLower()) {
          oldPanelName = (panel["panel"] as string).Replace("'", "");
        }
      }
      if (!String.IsNullOrEmpty(oldPanelName) && !String.IsNullOrEmpty(fedFrom)) {
        PanelUserControl fedFromUserControl = (PanelUserControl)findUserControl(fedFrom);
        fedFromUserControl.UpdateSubpanelName("PANEL " + oldPanelName, "PANEL " + input);
      }
      PanelUserControl userControl = (PanelUserControl)findUserControl(input);
      userControl.UpdateSubpanelFedFrom();
      SavePanelDataToLocalJsonFile();
    }

    private void MAINFORM_DEACTIVATE(object sender, EventArgs e) {
      foreach (PanelUserControl userControl in userControls) {
        DataGridView panelGrid = userControl.retrieve_panelGrid();
        panelGrid.ClearSelection();
      }
    }

    private void NEW_PANEL_BUTTON_Click(object sender, EventArgs e) {
      this.newPanelForm.ShowDialog();
    }

    private void CREATE_ALL_PANELS_BUTTON_Click(object sender, EventArgs e) {
      List<PanelUserControl> userControls = retrieve_userControls();
      List<Dictionary<string, object>> panels = new List<Dictionary<string, object>>();

      foreach (PanelUserControl userControl in userControls) {
        Dictionary<string, object> panelData = userControl.retrieve_data_from_modal();
        panels.Add(panelData);
      }

      using (
          DocumentLock docLock =
              Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
      ) {
        Close();
        myCommandsInstance.Create_Panels(panels);

        Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.WindowState = Autodesk.AutoCAD.Windows.Window.State.Maximized;
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Window.Focus();
      }
    }

    private void MAINFORM_SHOWN(object sender, EventArgs e) {
      // Check if the userControls list is empty
      if (this.userControls.Count == 0) {
        // If empty, show newPanelForm as a modal dialog
        newPanelForm.ShowDialog(); // or use appropriate method to show it as modal
      }
    }

    private void HELP_BUTTON_Click(object sender, EventArgs e) {
      HelpForm helpForm = new HelpForm();

      helpForm.ShowDialog();
    }

    private void SAVE_BUTTON_Click(object sender, EventArgs e) {
      SavePanelDataToLocalJsonFile();
    }

    private void MAINFORM_KEYDOWN(object sender, KeyEventArgs e) {
      if (e.Control && e.KeyCode == Keys.S) {
        SavePanelDataToLocalJsonFile();
      }
    }

    private void LOAD_BUTTON_Click(object sender, EventArgs e) {
      Close();
      // Prompt the user to select a JSON file
      OpenFileDialog openFileDialog = new OpenFileDialog {
        Filter = "JSON files (*.json)|*.json",
        Title = "Select a JSON file"
      };

      if (openFileDialog.ShowDialog() == DialogResult.OK) {
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

    private void DUPLICATE_PANEL_BUTTON_Click(object sender, EventArgs e) {
      DuplicatePanel();
    }

    private void LOAD_CALCULATIONS_BUTTON_Click(object sender, EventArgs e) {
      using (DocumentLock docLock = this.acDoc.LockDocument()) {
        CreateLoadCalculationsTable(this.userControls);
      }
    }

    public static void CreateLoadCalculationsTable(List<PanelUserControl> userControls) {
      Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;

      // Collect all subpanel names
      HashSet<string> subpanelNames = new HashSet<string>();
      foreach (var userControl in userControls) {
        subpanelNames.UnionWith(userControl.GetSubPanels());
      }

      using (DocumentLock docLock = doc.LockDocument()) {
        using (Transaction tr = db.TransactionManager.StartTransaction()) {
          try {
            BlockTableRecord currentSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
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
            for (int row = 0; row < rowCount; row++) {
              table.Rows[row].Height = 0.75;
              for (int col = 0; col < 3; col++) {
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

            foreach (var userControl in userControls) {
              string panelName = userControl.GetPanelName();
              string newOrExisting = userControl.GetNewOrExisting();
              if (!subpanelNames.Contains(panelName)) {
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
          catch (System.Exception ex) {
            ed.WriteMessage($"\nError creating load calculations table: {ex.Message}");
            tr.Abort();
          }
        }
      }
    }

    private static ObjectId CreateOrGetTextStyle(Database db, Transaction tr, string styleName) {
      TextStyleTable textStyleTable = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);

      if (!textStyleTable.Has(styleName)) {
        using (TextStyleTableRecord textStyle = new TextStyleTableRecord()) {
          textStyle.Name = styleName;
          textStyle.Font = new FontDescriptor(styleName, false, false, 0, 0);

          textStyleTable.UpgradeOpen();
          ObjectId textStyleId = textStyleTable.Add(textStyle);
          tr.AddNewlyCreatedDBObject(textStyle, true);

          return textStyleId;
        }
      }
      else {
        return textStyleTable[styleName];
      }
    }

    public void UpdateLCLLML() {
      Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;
      Editor ed = doc.Editor;
      LCLLMLManager manager = new LCLLMLManager();

      // First pass: Collect initial data
      foreach (PanelUserControl userControl in this.userControls) {
        LCLLMLObject obj = new LCLLMLObject(userControl.Name.Replace("'", ""));
        var LCLOverride = (int)userControl.GetLCLOverride();
        var LMLOverride = (int)userControl.GetLMLOverride();
        obj.LCLOVERRIDE = LCLOverride != 0;
        obj.LMLOVERRIDE = LMLOverride != 0;
        obj.LCL = (LCLOverride != 0) ? LCLOverride : (int)Math.Round(userControl.CalculateWattageSum("LCL"));
        obj.LML = (LMLOverride != 0) ? LMLOverride : (int)Math.Round(userControl.StoreItemsAndWattage("LML"));
        obj.Subpanels = userControl.GetSubPanels();
        manager.List.Add(obj);
      }

      // Second pass: Calculate final LCL and LML values
      foreach (var panel in manager.List) {
        CalculateLCL(panel, manager.List);
        CalculateLML(panel, manager.List);
      }

      // Third pass: Update user controls with calculated values
      foreach (PanelUserControl userControl in this.userControls) {
        var panelObj = manager.List.Find(p => p.PanelName == userControl.Name.Replace("'", ""));
        if (panelObj != null) {
          userControl.UpdateLCLLMLLabels(panelObj.LCL, panelObj.LML);
        }
      }
    }

    private void CalculateLCL(LCLLMLObject panel, List<LCLLMLObject> allPanels) {
      if (panel.LCLOVERRIDE) return;

      int totalLCL = panel.LCL;
      foreach (var subpanelName in panel.Subpanels) {
        var subpanel = allPanels.Find(p => p.PanelName == subpanelName);
        if (subpanel != null) {
          totalLCL += subpanel.LCL;
        }
      }
      panel.LCL = totalLCL;
    }

    private void CalculateLML(LCLLMLObject panel, List<LCLLMLObject> allPanels) {
      if (panel.LMLOVERRIDE) return;

      int maxLML = panel.LML;
      maxLML = RecursiveCalculateLML(panel, allPanels, maxLML);
      panel.LML = maxLML;
    }

    private int RecursiveCalculateLML(LCLLMLObject panel, List<LCLLMLObject> allPanels, int currentMax) {
      foreach (var subpanelName in panel.Subpanels) {
        var subpanel = allPanels.Find(p => p.PanelName == subpanelName);
        if (subpanel != null && !subpanel.LMLOVERRIDE) {
          currentMax = Math.Max(currentMax, subpanel.LML);
          currentMax = RecursiveCalculateLML(subpanel, allPanels, currentMax);
        }
      }
      return currentMax;
    }
  }

  public class LCLLMLObject {
    public int LCL { get; set; }
    public int LML { get; set; }
    public bool LCLOVERRIDE { get; set; }
    public bool LMLOVERRIDE { get; set; }
    public List<string> Subpanels { get; set; }
    public string PanelName { get; }

    public LCLLMLObject(string panelName) {
      LCL = 0;
      LML = 0;
      Subpanels = new List<string>();
      PanelName = panelName;
    }
  }

  public class LCLLMLManager {
    public List<LCLLMLObject> List { get; set; }

    public LCLLMLManager() {
      List = new List<LCLLMLObject>();
    }
  }
}