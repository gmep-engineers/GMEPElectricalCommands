﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using Newtonsoft.Json;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System.Linq;

namespace ElectricalCommands
{
  public partial class MainForm : Form
  {
    private PanelCommands myCommandsInstance;
    private NEWPANELFORM newPanelForm;
    private List<UserInterface> userControls;

    public MainForm(PanelCommands myCommands)
    {
      InitializeComponent();
      this.myCommandsInstance = myCommands;
      this.newPanelForm = new NEWPANELFORM(this);
      this.userControls = new List<UserInterface>();
      this.Shown += MAINFORM_SHOWN;
      this.FormClosing += MAINFORM_CLOSING;
      this.KeyPreview = true;
      this.KeyDown += new KeyEventHandler(MAINFORM_KEYDOWN);
      this.Deactivate += MAINFORM_DEACTIVATE;
      Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      acDoc.BeginDocumentClose -= new DocumentBeginCloseEventHandler(docBeginDocClose);
      acDoc.BeginDocumentClose += new DocumentBeginCloseEventHandler(docBeginDocClose);
    }

    private void docBeginDocClose(object sender, DocumentBeginCloseEventArgs e)
    {
      save_panel_to_autocad_document();
      Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      string fileName = acDoc.Name;
      acDoc.Database.SaveAs(fileName, true, DwgVersion.Current, acDoc.Database.SecurityParameters);
    }

    public List<UserInterface> retrieve_userControls()
    {
      return this.userControls;
    }

    public UserControl findUserControl(string panelName)
    {
      foreach (UserInterface userControl in userControls)
      {
        string userControlName = userControl.Name.Replace("'", "");
        userControlName = userControlName.Replace(" ", "");
        userControlName = userControlName.Replace("-", "");
        userControlName = userControlName.Replace("PANEL", "");

        panelName = panelName.Replace("'", "");
        panelName = panelName.Replace(" ", "");
        panelName = panelName.Replace("-", "");
        panelName = panelName.Replace("PANEL", "");

        if (userControlName.ToLower() == panelName.ToLower())
        {
          return userControl;
        }
      }

      return null;
    }

    public void initialize_modal()
    {
      PANEL_TABS.TabPages.Clear();

      List<Dictionary<string, object>> panelStorage = retrieve_saved_panel_data();

      if (panelStorage.Count == 0)
      {
        return;
      }
      else
      {
        set_up_cell_values_from_panel_data(panelStorage);
        set_up_tags_from_panel_data(panelStorage);
      }
    }

    public static void put_in_json_file(object thing)
    {
      // json write the panel data to the desktop
      string json = JsonConvert.SerializeObject(thing, Formatting.Indented);
      string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

      var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      //var ed = doc.Editor;

      //PromptResult pr = ed.GetString("\nEnter a file name: ");

      //string baseFileName = pr.StringResult;

      string baseFileName = "Test";

      if (string.IsNullOrEmpty(baseFileName))
      {
        baseFileName = "panel_data";
      }
      string extension = ".json";
      string path = Path.Combine(desktopPath, baseFileName + extension);

      int count = 1;
      while (File.Exists(path))
      {
        string tempFileName = string.Format("{0}({1})", baseFileName, count++);
        path = Path.Combine(desktopPath, tempFileName + extension);
      }

      File.WriteAllText(path, json);
    }

    private void set_up_cell_values_from_panel_data(List<Dictionary<string, object>> panelStorage)
    {
      foreach (Dictionary<string, object> panel in panelStorage)
      {
        string panelName = panel["panel"].ToString();
        bool is3PH = panel.ContainsKey("phase_c_left");
        UserInterface userControl1 = create_new_panel_tab(panelName, is3PH);
        userControl1.clear_modal_and_remove_rows(panel);
        userControl1.populate_modal_with_panel_data(panel);
        var notes = JsonConvert.DeserializeObject<List<string>>(panel["notes"].ToString());
        userControl1.update_notes_storage(notes);
      }
    }

    private void set_up_tags_from_panel_data(List<Dictionary<string, object>> panelStorage)
    {
      foreach (Dictionary<string, object> panel in panelStorage)
      {
        string panelName = panel["panel"].ToString();
        UserInterface userControl1 = (UserInterface)findUserControl(panelName);
        if (userControl1 == null)
        {
          continue;
        }
        DataGridView panelGrid = userControl1.retrieve_panelGrid();
        foreach (DataGridViewRow row in panelGrid.Rows)
        {
          int rowIndex = row.Index;
          var tagNames = new Dictionary<string, int>();
          if (panel.ContainsKey("phase_c_left"))
          {
            tagNames = new Dictionary<string, int>()
            {
              {"phase_a_left_tag", 1},
              {"phase_b_left_tag", 2},
              {"phase_c_left_tag", 3},
              {"phase_a_right_tag", 8},
              {"phase_b_right_tag", 9},
              {"phase_c_right_tag", 10},
              {"description_left_tags", 0},
              {"description_right_tags", 11}
            };
          }
          else
          {
            tagNames = new Dictionary<string, int>()
            {
              {"phase_a_left_tag", 1},
              {"phase_b_left_tag", 2},
              {"phase_a_right_tag", 7},
              {"phase_b_right_tag", 8},
              {"description_left_tags", 0},
              {"description_right_tags", 9}
            };
          }

          foreach (var tagName in tagNames)
          {
            if (panel.ContainsKey(tagName.Key))
            {
              set_cell_value(panel, tagName.Key, rowIndex, tagName.Value, row);
            }
          }
        }
        userControl1.update_cell_background_color();
        userControl1.recalculate_breakers();
        userControl1.calculate_lcl_otherload_panelload_feederamps();
      }
    }

    private void set_cell_value(Dictionary<string, object> panel, string key, int rowIndex, int cellIndex, DataGridViewRow row)
    {
      string tag = panel[key].ToString();
      List<string> tagList = JsonConvert.DeserializeObject<List<string>>(tag);
      string tagValue = tagList[rowIndex];
      if (tagValue != "")
      {
        if (key.Contains("phase"))
        {
          row.Cells[cellIndex].Value = tagValue;
        }
        else if (key.Contains("description"))
        {
          row.Cells[cellIndex].Tag = tagValue;
        }
      }
    }

    internal void delete_panel(UserInterface userControl1)
    {
      DialogResult dialogResult = MessageBox.Show("Are you sure you want to delete this panel?", "Delete Panel", MessageBoxButtons.YesNo);
      if (dialogResult == DialogResult.Yes)
      {
        this.userControls.Remove(userControl1);
        PANEL_TABS.TabPages.Remove(userControl1.Parent as TabPage);
        using (DocumentLock docLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
        {
          remove_panel_from_storage(userControl1);
        }
      }
    }

    private void remove_panel_from_storage(UserInterface userControl1)
    {
      var panelData = retrieve_saved_panel_data();

      foreach (Dictionary<string, object> panel in panelData)
      {
        var panelName = panel["panel"].ToString().Replace("\'", "").Replace("`", "");
        var userControlName = userControl1.Name.Replace("\'", "").Replace("`", "");
        if (panelName == userControlName)
        {
          panelData.Remove(panel);
          break;
        }
      }

      store_data_in_autocad_file(panelData);
    }

    internal bool panel_name_exists(string panelName)
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

    public void add_usercontrol_to_new_tab(UserControl control, TabPage tabPage)
    {
      // Set the user control location and size if needed
      control.Location = new Point(0, 0); // Top-left corner of the tab page
      control.Dock = DockStyle.Fill; // If you want to dock it to fill the tab

      // Add the user control to the controls of the tab page
      tabPage.Controls.Add(control);
    }

    public UserInterface create_new_panel_tab(string tabName, bool is3PH)
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
      UserInterface userControl1 = new UserInterface(this.myCommandsInstance, this, this.newPanelForm, tabName, is3PH);

      // Add the UserControl to the list of UserControls
      this.userControls.Add(userControl1);

      // Call the method to add the UserControl to the new tab
      add_usercontrol_to_new_tab(userControl1, newTabPage);

      return userControl1;
    }

    //public void store_data_in_autocad_file(List<Dictionary<string, object>> saveData)
    //{
    //  Autodesk.AutoCAD.ApplicationServices.Document acDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
    //  Autodesk.AutoCAD.DatabaseServices.Database acCurDb = acDoc.Database;
    //  string jsonDataKey = "JsonSaveData";

    //  using (Autodesk.AutoCAD.DatabaseServices.Transaction tr = acCurDb.TransactionManager.StartTransaction())
    //  {
    //    acCurDb.DisableUndoRecording(true);
    //    Autodesk.AutoCAD.DatabaseServices.DBDictionary nod = (Autodesk.AutoCAD.DatabaseServices.DBDictionary)tr.GetObject(acCurDb.NamedObjectsDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

    //    Autodesk.AutoCAD.DatabaseServices.DBDictionary userDict;
    //    if (nod.Contains(jsonDataKey))
    //    {
    //      The dictionary already exists, so we just need to open it for write
    //     userDict = (Autodesk.AutoCAD.DatabaseServices.DBDictionary)tr.GetObject(nod.GetAt(jsonDataKey), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

    //      Clear the dictionary

    //     userDict.UpgradeOpen();
    //      userDict.Erase(true);
    //      userDict.DowngradeOpen();

    //      Create a new dictionary
    //      userDict = new Autodesk.AutoCAD.DatabaseServices.DBDictionary();
    //      nod.UpgradeOpen();
    //      nod.SetAt(jsonDataKey, userDict);
    //      tr.AddNewlyCreatedDBObject(userDict, true);
    //    }
    //    else
    //    {
    //      The dictionary doesn't exist, so we create a new one and add it to the NOD
    //      userDict = new Autodesk.AutoCAD.DatabaseServices.DBDictionary();
    //      nod.UpgradeOpen();
    //      nod.SetAt(jsonDataKey, userDict);
    //      tr.AddNewlyCreatedDBObject(userDict, true);
    //    }

    //    Now let's update or create the Xrecord for each panel
    //    for (int i = 0; i < saveData.Count; i++)
    //    {
    //      string panelKey = "PanelData" + i.ToString("D3");
    //      Autodesk.AutoCAD.DatabaseServices.Xrecord xRecord;
    //      if (userDict.Contains(panelKey))
    //      {
    //        The Xrecord exists, open it for write to update

    //       xRecord = (Autodesk.AutoCAD.DatabaseServices.Xrecord)tr.GetObject(userDict.GetAt(panelKey), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);
    //      }
    //      else
    //      {
    //        The Xrecord does not exist, create a new one
    //        xRecord = new Autodesk.AutoCAD.DatabaseServices.Xrecord();
    //        userDict.SetAt(panelKey, xRecord);
    //        tr.AddNewlyCreatedDBObject(xRecord, true);
    //      }

    //      Update the Xrecord data
    //      Autodesk.AutoCAD.DatabaseServices.ResultBuffer rb = new Autodesk.AutoCAD.DatabaseServices.ResultBuffer(new Autodesk.AutoCAD.DatabaseServices.TypedValue((int)Autodesk.AutoCAD.DatabaseServices.DxfCode.Text, JsonConvert.SerializeObject(saveData[i], Formatting.Indented)));
    //      xRecord.Data = new Autodesk.AutoCAD.DatabaseServices.ResultBuffer();
    //      xRecord.Data = rb;
    //    }

    //    tr.Commit();
    //    acCurDb.DisableUndoRecording(false);
    //  }
    //}

    public List<Dictionary<string, object>> retrieve_saved_panel_data()
    {
      List<Dictionary<string, object>> allPanelData = new List<Dictionary<string, object>>();

      Autodesk.AutoCAD.ApplicationServices.Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      string acDocPath = Path.GetDirectoryName(acDoc.Name);
      string panelSavesDirectory = Path.Combine(acDocPath, "panel saves");

      // Check if the directory exists
      if (Directory.Exists(panelSavesDirectory))
      {
        // Get all JSON files in the directory
        string[] jsonFiles = Directory.GetFiles(panelSavesDirectory, "*.json");

        // If there are any JSON files, find the most recent one
        if (jsonFiles.Length > 0)
        {
          string mostRecentJsonFile = jsonFiles.OrderByDescending(f => File.GetLastWriteTime(f)).First();

          // Read the JSON data from the file
          string jsonData = File.ReadAllText(mostRecentJsonFile);

          // Deserialize the JSON data to a list of dictionaries
          allPanelData = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(jsonData);
        }
      }

      return allPanelData;
    }

    //public List<Dictionary<string, object>> retrieve_saved_panel_data()
    //{
    //  List<Dictionary<string, object>> allPanelData = new List<Dictionary<string, object>>();

    //  Autodesk.AutoCAD.ApplicationServices.Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
    //  Autodesk.AutoCAD.DatabaseServices.Database acCurDb = acDoc.Database;
    //  string jsonDataKey = "JsonSaveData";

    //  using (Autodesk.AutoCAD.DatabaseServices.Transaction tr = acCurDb.TransactionManager.StartTransaction())
    //  {
    //    Autodesk.AutoCAD.DatabaseServices.DBDictionary nod = (Autodesk.AutoCAD.DatabaseServices.DBDictionary)tr.GetObject(acCurDb.NamedObjectsDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

    //    if (nod.Contains(jsonDataKey))
    //    {
    //      Autodesk.AutoCAD.DatabaseServices.DBDictionary userDict = (Autodesk.AutoCAD.DatabaseServices.DBDictionary)tr.GetObject(nod.GetAt(jsonDataKey), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

    //      // Iterate over all XRecords in the user dictionary
    //      foreach (var entry in userDict)
    //      {
    //        Autodesk.AutoCAD.DatabaseServices.Xrecord xRecord = (Autodesk.AutoCAD.DatabaseServices.Xrecord)tr.GetObject(entry.Value, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
    //        Autodesk.AutoCAD.DatabaseServices.ResultBuffer rb = xRecord.Data;
    //        if (rb != null)
    //        {
    //          foreach (Autodesk.AutoCAD.DatabaseServices.TypedValue tv in rb)
    //          {
    //            if (tv.TypeCode == (int)Autodesk.AutoCAD.DatabaseServices.DxfCode.Text)
    //            {
    //              Dictionary<string, object> panelData = JsonConvert.DeserializeObject<Dictionary<string, object>>(tv.Value.ToString());
    //              allPanelData.Add(panelData);
    //            }
    //          }
    //        }
    //      }
    //    }
    //  }
    //  return allPanelData;
    //}

    public void save_panel_to_autocad_document()
    {
      List<Dictionary<string, object>> panelStorage = new List<Dictionary<string, object>>();

      var acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

      if (acDoc != null)
      {
        using (DocumentLock docLock = acDoc.LockDocument())
        {
          foreach (UserInterface userControl in this.userControls)
          {
            panelStorage.Add(userControl.retrieve_data_from_modal());
          }

          store_data_in_autocad_file(panelStorage);
        }
      }
    }

    private void MAINFORM_CLOSING(object sender, FormClosingEventArgs e)
    {
      Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      acDoc.BeginDocumentClose -= new DocumentBeginCloseEventHandler(docBeginDocClose);
      save_panel_to_autocad_document();
    }

    public void PANEL_NAME_INPUT_TextChanged(object sender, EventArgs e, string input)
    {
      // Get the selected tab index
      int selectedIndex = PANEL_TABS.SelectedIndex;

      // Check if there is a selected tab
      if (selectedIndex >= 0)
      {
        // Rename the current selected tab
        PANEL_TABS.TabPages[selectedIndex].Text = "Panel " + input;
      }
    }

    private void MAINFORM_DEACTIVATE(object sender, EventArgs e)
    {
      foreach (UserInterface userControl in userControls)
      {
        DataGridView panelGrid = userControl.retrieve_panelGrid();
        panelGrid.ClearSelection();
      }
    }

    private void NEW_PANEL_BUTTON_Click(object sender, EventArgs e)
    {
      this.newPanelForm.ShowDialog();
    }

    private void CREATE_ALL_PANELS_BUTTON_Click(object sender, EventArgs e)
    {
      List<UserInterface> userControls = retrieve_userControls();
      List<Dictionary<string, object>> panels = new List<Dictionary<string, object>>();

      foreach (UserInterface userControl in userControls)
      {
        Dictionary<string, object> panelData = userControl.retrieve_data_from_modal();
        panels.Add(panelData);
      }

      using (DocumentLock docLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
      {
        Close();
        myCommandsInstance.Create_Panels(panels);
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
      Help helpForm = new Help();

      helpForm.ShowDialog();
    }

    private void SAVE_BUTTON_Click(object sender, EventArgs e)
    {
      save_panel_to_autocad_document();
    }

    private void MAINFORM_KEYDOWN(object sender, KeyEventArgs e)
    {
      if (e.Control && e.KeyCode == Keys.S)
      {
        save_panel_to_autocad_document();
      }
    }

    private void LOAD_BUTTON_Click(object sender, EventArgs e)
    {
      Close();
      // Prompt the user to select a JSON file
      OpenFileDialog openFileDialog = new OpenFileDialog
      {
        Filter = "JSON files (*.json)|*.json",
        Title = "Select a JSON file"
      };

      if (openFileDialog.ShowDialog() == DialogResult.OK)
      {
        // Read the JSON data from the file
        string jsonData = File.ReadAllText(openFileDialog.FileName);

        // Deserialize the JSON data to a list of dictionaries
        List<Dictionary<string, object>> panelData = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(jsonData);

        // Save the panel data
        store_data_in_autocad_file(panelData);
      }
    }

    public void store_data_in_autocad_file(List<Dictionary<string, object>> saveData)
    {
      Autodesk.AutoCAD.ApplicationServices.Document acDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
      string acDocPath = Path.GetDirectoryName(acDoc.Name);
      string panelSavesDirectory = Path.Combine(acDocPath, "panel saves");

      // Create the directory if it doesn't exist
      if (!Directory.Exists(panelSavesDirectory))
      {
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
  }
}