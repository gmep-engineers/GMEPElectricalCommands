using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsSystem;
using DocumentFormat.OpenXml.Presentation;
using Emgu.CV.ImgHash;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml.Packaging.Ionic.Zlib;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ElectricalCommands {

  public partial class PanelUserControl : UserControl {
    private PanelCommands myCommandsInstance;
    private MainForm mainForm;
    private NewPanelForm newPanelForm;
    private NoteForm notesForm;
    private List<string> notesStorage = new List<string>();
    private List<DataGridViewCell> selectedCells;
    private List<string> defaultNotes;
    private object oldValue;
    private bool isLoading;
    private bool contains3PhEquip;

    public PanelUserControl(
      PanelCommands myCommands,
      MainForm mainForm,
      NewPanelForm newPanelForm,
      string tabName,
      bool is3Ph = false
    ) {
      InitializeComponent();
      this.isLoading = true;
      myCommandsInstance = myCommands;
      this.mainForm = mainForm;
      this.newPanelForm = newPanelForm;
      this.Name = tabName;
      this.notesStorage = new List<string>();
      this.is3Ph = is3Ph;
      this.contains3PhEquip = false;

      INFO_LABEL.Text = "";

      ListenForNewRows();
      AddOrRemovePanelGridColumns(is3Ph);
      RemoveColumnHeaderSorting();

      ChangeSizeOfPhaseColumns(is3Ph);
      AddPhaseSumColumn(is3Ph);

      PANEL_NAME_INPUT.TextChanged += new EventHandler(this.PANEL_NAME_INPUT_TextChanged);
      PANEL_GRID.CellValueChanged += new DataGridViewCellEventHandler(this.PANEL_GRID_CellValueChangedLink);
      PANEL_GRID.Rows.AddCopies(0, 21);
      PANEL_GRID.AllowUserToAddRows = false;

      AddRowsToDatagrid();
      SetDefaultFormValues(tabName);
      DeselectCells();
    }

    private void TogglePrefixInSelectedCells(string prefix) {
      bool allCellsEmptyOrWithPrefix = true;
      List<DataGridViewCell> cellsToUpdate = new List<DataGridViewCell>();

      // First pass: Check if all selected cells are empty or have the prefix
      foreach (DataGridViewCell cell in PANEL_GRID.SelectedCells) {
        if (cell.Value != null && PANEL_GRID.Columns[cell.ColumnIndex].Name.ToLower().Contains("description")) {
          string cellValue = cell.Value.ToString();
          bool hasPrefixAtStart = cellValue.StartsWith(prefix);
          bool hasPrefixAfterSemicolon = cellValue.Contains(";" + prefix);

          if (string.IsNullOrEmpty(cellValue) || hasPrefixAtStart || (cellValue.Contains(";") && hasPrefixAfterSemicolon)) {
            // Cell is empty or has the prefix, do nothing
          }
          else {
            allCellsEmptyOrWithPrefix = false;
          }
          cellsToUpdate.Add(cell);
        }
      }

      // Second pass: Update cell values based on the check
      foreach (DataGridViewCell cell in cellsToUpdate) {
        if (cell.Value != null) {
          string cellValue = cell.Value.ToString();
          if (allCellsEmptyOrWithPrefix) {
            // Remove the prefix at the beginning
            if (cellValue.StartsWith(prefix)) {
              cellValue = cellValue.Substring(prefix.Length);
            }
            // Remove the prefix after the semicolon if it exists
            if (cellValue.Contains(";" + prefix)) {
              cellValue = cellValue.Replace(";" + prefix, ";");
            }
            cell.Value = cellValue;
          }
          else {
            // Add the prefix at the beginning if not present and cell is not empty
            if (!cellValue.StartsWith(prefix) && !string.IsNullOrEmpty(cellValue)) {
              cellValue = prefix + cellValue;
            }
            // Add the prefix after the semicolon if it exists and not already present
            if (cellValue.Contains(";") && !cellValue.Contains(";" + prefix) && !string.IsNullOrEmpty(cellValue)) {
              int semicolonIndex = cellValue.IndexOf(";");
              cellValue = cellValue.Insert(semicolonIndex + 1, prefix);
            }
            cell.Value = cellValue;
          }
        }
      }
    }

    private void EXISTING_BUTTON_Click(object sender, EventArgs e) {
      TogglePrefixInSelectedCells("(E)");
    }

    private void RELOCATE_BUTTON_Click(object sender, EventArgs e) {
      TogglePrefixInSelectedCells("(R)");
    }

    public List<string> GetNotesStorage() {
      return this.notesStorage;
    }

    private void AddRowsToDatagrid() {
      PHASE_SUM_GRID.Rows.Add("0", "0");
      TOTAL_VA_GRID.Rows.Add("0");
      PANEL_LOAD_GRID.Rows.Add("0");
      FEEDER_AMP_GRID.Rows.Add("0");
    }

    private void DeselectCells() {
      PHASE_SUM_GRID.DefaultCellStyle.SelectionBackColor = PHASE_SUM_GRID
        .DefaultCellStyle
        .BackColor;
      PHASE_SUM_GRID.DefaultCellStyle.SelectionForeColor = PHASE_SUM_GRID
        .DefaultCellStyle
        .ForeColor;
      TOTAL_VA_GRID.DefaultCellStyle.SelectionBackColor = TOTAL_VA_GRID.DefaultCellStyle.BackColor;
      TOTAL_VA_GRID.DefaultCellStyle.SelectionForeColor = TOTAL_VA_GRID.DefaultCellStyle.ForeColor;
      PANEL_LOAD_GRID.DefaultCellStyle.SelectionBackColor = PANEL_LOAD_GRID
        .DefaultCellStyle
        .BackColor;
      PANEL_LOAD_GRID.DefaultCellStyle.SelectionForeColor = PANEL_LOAD_GRID
        .DefaultCellStyle
        .ForeColor;
      FEEDER_AMP_GRID.DefaultCellStyle.SelectionBackColor = FEEDER_AMP_GRID
        .DefaultCellStyle
        .BackColor;
      FEEDER_AMP_GRID.DefaultCellStyle.SelectionForeColor = FEEDER_AMP_GRID
        .DefaultCellStyle
        .ForeColor;
      PANEL_GRID.ClearSelection();
    }

    private void SetDefaultFormValues(string tabName) {
      // Textboxes
      PANEL_NAME_INPUT.Text = tabName;
      PANEL_LOCATION_INPUT.Text = "ELECTRIC ROOM";
      MAIN_INPUT.Text = "M.L.O.";
      BUS_RATING_INPUT.Text = "100";

      // Comboboxes
      STATUS_COMBOBOX.SelectedIndex = 0;
      MOUNTING_COMBOBOX.SelectedIndex = 0;
      if (PHASE_SUM_GRID.ColumnCount > 2) {
        WIRE_COMBOBOX.SelectedIndex = 1;
        PHASE_COMBOBOX.SelectedIndex = 1;
      }
      else {
        WIRE_COMBOBOX.SelectedIndex = 0;
        PHASE_COMBOBOX.SelectedIndex = 0;
      }
      LINE_VOLTAGE_COMBOBOX.SelectedIndex = 0;
      PHASE_VOLTAGE_COMBOBOX.SelectedIndex = 0;

      // Datagrids
      PHASE_SUM_GRID.Rows[0].Cells[0].Value = "0";
      PHASE_SUM_GRID.Rows[0].Cells[1].Value = "0";
      TOTAL_VA_GRID.Rows[0].Cells[0].Value = "0";
      PANEL_LOAD_GRID.Rows[0].Cells[0].Value = "0";
      FEEDER_AMP_GRID.Rows[0].Cells[0].Value = "0";

      if (PHASE_SUM_GRID.ColumnCount > 2)
        PHASE_SUM_GRID.Rows[0].Cells[2].Value = "0";
    }

    private void RemoveColumnHeaderSorting() {
      foreach (DataGridViewColumn column in PANEL_GRID.Columns) {
        column.SortMode = DataGridViewColumnSortMode.NotSortable;
      }
    }

    private void ListenForNewRows() {
      PANEL_GRID.RowsAdded += new DataGridViewRowsAddedEventHandler(PANEL_GRID_RowsAdded);
    }

    public static double SafeConvertToDouble(string value) {
      if (double.TryParse(value, out double result)) {
        return result;
      }
      return 0;
    }

    public Dictionary<string, object> RetrieveDataFromModal() {
      // Create a new panel
      Dictionary<string, object> panel = new Dictionary<string, object>();

      if (String.IsNullOrEmpty(id)) {
        id = System.Guid.NewGuid().ToString();
      }

      panel.Add("id", id);

      // Get the value from the main input
      string mainInput = MAIN_INPUT.Text.ToLower();

      if (
        !mainInput.Contains("mlo")
        && !mainInput.Contains("m.l.o")
        && !mainInput.Contains("m.l.o.")
      ) {
        if (mainInput.Contains("amp")) {
          mainInput = mainInput.Replace("amp", "AMP");
        }
        else if (mainInput.Contains("a")) {
          mainInput = mainInput.Replace("a", "AMP");
        }
        else if (mainInput.Contains(" ")) {
          mainInput = mainInput.Replace(" ", " AMP");
        }
        else {
          mainInput += " AMP";
        }
      }

      panel.Add("main", mainInput.ToUpper());

      string GetComboBoxValue(ComboBox comboBox) {
        if (comboBox.SelectedItem != null) {
          return comboBox.SelectedItem.ToString().ToUpper();
        }
        else if (!string.IsNullOrEmpty(comboBox.Text)) {
          return comboBox.Text.ToUpper();
        }
        else {
          return "";
        }
      }

      panel.Add("panel", "'" + PANEL_NAME_INPUT.Text.ToUpper() + "'");
      panel.Add("location", PANEL_LOCATION_INPUT.Text.ToUpper());
      panel.Add("voltage1", GetComboBoxValue(PHASE_VOLTAGE_COMBOBOX));
      panel.Add("voltage2", GetComboBoxValue(LINE_VOLTAGE_COMBOBOX));
      panel.Add("phase", GetComboBoxValue(PHASE_COMBOBOX));
      panel.Add("wire", GetComboBoxValue(WIRE_COMBOBOX));
      panel.Add("mounting", GetComboBoxValue(MOUNTING_COMBOBOX));
      panel.Add("existing", GetComboBoxValue(STATUS_COMBOBOX));
      panel.Add("lcl_override", LCL_OVERRIDE.Checked);
      panel.Add("lml_override", LML_OVERRIDE.Checked);
      panel.Add("distribution_section", DISTRIBUTION_SECTION_CHECKBOX.Checked);

      panel.Add(
        "subtotal_a",
        Math.Round(Convert.ToDouble(PHASE_SUM_GRID.Rows[0].Cells[0].Value.ToString().ToUpper()))
          .ToString()
      );
      panel.Add(
        "subtotal_b",
        Math.Round(Convert.ToDouble(PHASE_SUM_GRID.Rows[0].Cells[1].Value.ToString().ToUpper()))
          .ToString()
      );

      if (PHASE_SUM_GRID.Columns.Count > 2) {
        panel.Add(
          "subtotal_c",
          Math.Round(Convert.ToDouble(PHASE_SUM_GRID.Rows[0].Cells[2].Value.ToString().ToUpper()))
            .ToString()
        );
      }
      else {
        panel.Add("subtotal_c", "0");
      }
      panel.Add("total_va", TOTAL_VA_GRID.Rows[0].Cells[0].Value.ToString().ToUpper());
      double lcl = SafeConvertToDouble(LCL.Text);
      double lcl125 = Math.Round(lcl * 1.25, 0);
      panel.Add("lcl", lcl);
      panel.Add("lcl125", lcl125);
      double lml = SafeConvertToDouble(LML.Text);
      double lml125 = Math.Round(lml * 1.25, 0);
      panel.Add("lml", lml);
      panel.Add("lml125", lml125);
      double safetyFactor = 1.0;
      if (Regex.IsMatch(SAFETY_FACTOR_TEXTBOX.Text, @"^\d*\.?\d*$")) {
        safetyFactor = SafeConvertToDouble(SAFETY_FACTOR_TEXTBOX.Text);
      }
      else {
        SAFETY_FACTOR_CHECKBOX.Checked = false;
      }
      panel.Add("using_safety_factor", SAFETY_FACTOR_CHECKBOX.Checked);
      panel.Add("safety_factor", safetyFactor);
      if (SAFETY_FACTOR_CHECKBOX.Enabled && SAFETY_FACTOR_CHECKBOX.Checked) {
        double feederAmps = Convert.ToDouble(FEEDER_AMP_GRID.Rows[0].Cells[0].Value.ToString()) / safetyFactor;
        double kva = Convert.ToDouble(PANEL_LOAD_GRID.Rows[0].Cells[0].Value.ToString()) / safetyFactor;
        panel.Add("feeder_amps", feederAmps.ToString());
        panel.Add("kva", Math.Round(kva, 1).ToString());
      }
      else {
        panel.Add("kva", PANEL_LOAD_GRID.Rows[0].Cells[0].Value.ToString().ToUpper());
        panel.Add("feeder_amps", FEEDER_AMP_GRID.Rows[0].Cells[0].Value.ToString());
      }
      panel.Add("custom_title", CUSTOM_TITLE_TEXT.Text.ToUpper());
      panel.Add("fed_from", FED_FROM_TEXTBOX.Text.ToUpper());

      string busRatingInput = BUS_RATING_INPUT.Text.ToLower();

      if (busRatingInput.Contains("amp")) {
        busRatingInput = busRatingInput.Replace("amp", "A");
      }
      else if (busRatingInput.Contains("a")) {
        busRatingInput = busRatingInput.Replace("a", "A");
      }
      else if (busRatingInput.Contains(" ")) {
        busRatingInput = busRatingInput.Replace(" ", " A");
      }
      else {
        busRatingInput += "A";
      }

      panel.Add("bus_rating", busRatingInput.ToUpper());

      List<bool> description_left_highlights = new List<bool>();
      List<bool> description_right_highlights = new List<bool>();
      List<bool> breaker_left_highlights = new List<bool>();
      List<bool> breaker_right_highlights = new List<bool>();

      List<string> description_left = new List<string>();
      List<string> description_right = new List<string>();
      List<string> phase_a_left = new List<string>();
      List<string> phase_b_left = new List<string>();
      List<string> phase_a_right = new List<string>();
      List<string> phase_b_right = new List<string>();
      List<string> phase_c_left = new List<string>();
      List<string> phase_c_right = new List<string>();
      List<string> breaker_left = new List<string>();
      List<string> breaker_right = new List<string>();
      List<string> circuit_left = new List<string>();
      List<string> circuit_right = new List<string>();
      List<string> phase_a_left_tag = new List<string>();
      List<string> phase_b_left_tag = new List<string>();
      List<string> phase_a_right_tag = new List<string>();
      List<string> phase_b_right_tag = new List<string>();
      List<string> phase_c_left_tag = new List<string>();
      List<string> phase_c_right_tag = new List<string>();

      List<string> description_left_tags = new List<string>();
      List<string> description_right_tags = new List<string>();
      for (int i = 0; i < PANEL_GRID.Rows.Count; i++) {
        string descriptionLeftValue = "";
        if (
          string.IsNullOrEmpty(PANEL_GRID.Rows[i].Cells["description_left"].Value as string)
          && !string.IsNullOrEmpty(PANEL_GRID.Rows[i].Cells["breaker_left"].Value as string)
        ) {
          // check that the breaker value is both an integer and greater than 3
          var breakerValue = PANEL_GRID.Rows[i].Cells["breaker_left"].Value.ToString();
          int breakerValueInt;
          if (int.TryParse(breakerValue, out breakerValueInt) && breakerValueInt > 3) {
            descriptionLeftValue = "SPARE";
          }
        }
        else {
          descriptionLeftValue = string.IsNullOrWhiteSpace(
            PANEL_GRID.Rows[i].Cells["description_left"].Value as string
          )
            ? "SPACE"
            : PANEL_GRID
              .Rows[i]
              .Cells["description_left"]
              .Value.ToString()
              .ToUpper()
              .Replace("\r", "");
        }
        if (descriptionLeftValue.StartsWith("PANEL")) {
          foreach (PanelUserControl userControl in this.mainForm.RetrieveUserControls()) {
            TextBox panelName =
              userControl.Controls.Find("PANEL_NAME_INPUT", true).FirstOrDefault() as TextBox;
            if ("PANEL " + panelName.Text == descriptionLeftValue) {
              descriptionLeftValue = !String.IsNullOrEmpty(userControl.GetId()) ? userControl.GetId() : panelName.Text;
            }
          }
        }
        string descriptionRightValue = "";
        if (
          string.IsNullOrEmpty(PANEL_GRID.Rows[i].Cells["description_right"].Value as string)
          && !string.IsNullOrEmpty(PANEL_GRID.Rows[i].Cells["breaker_right"].Value as string)
        ) {
          // check that the breaker value is both an integer and greater than 3
          var breakerValue = PANEL_GRID.Rows[i].Cells["breaker_right"].Value.ToString();
          int breakerValueInt;
          if (int.TryParse(breakerValue, out breakerValueInt) && breakerValueInt > 3) {
            descriptionRightValue = "SPARE";
          }
        }
        else {
          descriptionRightValue = string.IsNullOrWhiteSpace(
            PANEL_GRID.Rows[i].Cells["description_right"].Value as string
          )
            ? "SPACE"
            : PANEL_GRID
              .Rows[i]
              .Cells["description_right"]
              .Value.ToString()
              .ToUpper()
              .Replace("\r", "");
        }
        if (descriptionRightValue.StartsWith("PANEL")) {
          foreach (PanelUserControl userControl in this.mainForm.RetrieveUserControls()) {
            TextBox panelName =
              userControl.Controls.Find("PANEL_NAME_INPUT", true).FirstOrDefault() as TextBox;
            if ("PANEL " + panelName.Text == descriptionRightValue) {
              descriptionRightValue = !String.IsNullOrEmpty(userControl.GetId()) ? userControl.GetId() : panelName.Text;
            }
          }
        }
        string breakerLeftValue =
          PANEL_GRID.Rows[i].Cells["breaker_left"].Value?.ToString().ToUpper().Replace("\r", "")
          ?? "";
        string breakerRightValue =
          PANEL_GRID.Rows[i].Cells["breaker_right"].Value?.ToString().ToUpper().Replace("\r", "")
          ?? "";
        string circuitRightValue =
          PANEL_GRID.Rows[i].Cells["circuit_right"].Value?.ToString().ToUpper().Replace("\r", "")
          ?? "";
        string circuitLeftValue =
          PANEL_GRID.Rows[i].Cells["circuit_left"].Value?.ToString().ToUpper().Replace("\r", "")
          ?? "";
        string phaseALeftTag = PANEL_GRID.Rows[i].Cells["phase_a_left"].Tag?.ToString() ?? "";
        string phaseBLeftTag = PANEL_GRID.Rows[i].Cells["phase_b_left"].Tag?.ToString() ?? "";
        string phaseARightTag = PANEL_GRID.Rows[i].Cells["phase_a_right"].Tag?.ToString() ?? "";
        string phaseBRightTag = PANEL_GRID.Rows[i].Cells["phase_b_right"].Tag?.ToString() ?? "";

        string phaseALeftValue = (
          PANEL_GRID
            .Rows[i]
            .Cells["phase_a_left"]
            .Value?.ToString()
            .Replace("\r", "")
            .Replace(" ", "") ?? "0"
        );
        phaseALeftValue =
          phaseALeftValue.Contains(";") || !Regex.IsMatch(phaseALeftValue, @"^-?\d+$")
            ? phaseALeftValue
            : Math.Round(Convert.ToDouble(phaseALeftValue)).ToString();

        string phaseBLeftValue = (
          PANEL_GRID
            .Rows[i]
            .Cells["phase_b_left"]
            .Value?.ToString()
            .Replace("\r", "")
            .Replace(" ", "") ?? "0"
        );
        phaseBLeftValue =
          phaseBLeftValue.Contains(";") || !Regex.IsMatch(phaseBLeftValue, @"^-?\d+$")
            ? phaseBLeftValue
            : Math.Round(Convert.ToDouble(phaseBLeftValue)).ToString();

        string phaseARightValue = (
          PANEL_GRID
            .Rows[i]
            .Cells["phase_a_right"]
            .Value?.ToString()
            .Replace("\r", "")
            .Replace(" ", "") ?? "0"
        );
        phaseARightValue =
          phaseARightValue.Contains(";") || !Regex.IsMatch(phaseARightValue, @"^-?\d+$")
            ? phaseARightValue
            : Math.Round(Convert.ToDouble(phaseARightValue)).ToString();

        string phaseBRightValue = (
          PANEL_GRID
            .Rows[i]
            .Cells["phase_b_right"]
            .Value?.ToString()
            .Replace("\r", "")
            .Replace(" ", "") ?? "0"
        );
        phaseBRightValue =
          phaseBRightValue.Contains(";") || !Regex.IsMatch(phaseBRightValue, @"^-?\d+$")
            ? phaseBRightValue
            : Math.Round(Convert.ToDouble(phaseBRightValue)).ToString();

        string phaseCLeftValue = "0";
        string phaseCRightValue = "0";
        string phaseCLeftTag = "";
        string phaseCRightTag = "";

        string descriptionLeftTag =
          PANEL_GRID.Rows[i].Cells["description_left"].Tag?.ToString() ?? "";
        string descriptionRightTag =
          PANEL_GRID.Rows[i].Cells["description_right"].Tag?.ToString() ?? "";

        if (PHASE_SUM_GRID.Columns.Count > 2) {
          phaseCLeftTag = PANEL_GRID.Rows[i].Cells["phase_c_left"].Tag?.ToString() ?? "";
          phaseCRightTag = PANEL_GRID.Rows[i].Cells["phase_c_right"].Tag?.ToString() ?? "";
          phaseCLeftValue = (
            PANEL_GRID
              .Rows[i]
              .Cells["phase_c_left"]
              .Value?.ToString()
              .Replace("\r", "")
              .Replace(" ", "") ?? "0"
          );
          phaseCLeftValue =
            phaseCLeftValue.Contains(";") || !Regex.IsMatch(phaseCLeftValue, @"^-?\d+$")
              ? phaseCLeftValue
              : Math.Round(Convert.ToDouble(phaseCLeftValue)).ToString();
          phaseCRightValue = (
            PANEL_GRID
              .Rows[i]
              .Cells["phase_c_right"]
              .Value?.ToString()
              .Replace("\r", "")
              .Replace(" ", "") ?? "0"
          );
          phaseCRightValue =
            phaseCRightValue.Contains(";") || !Regex.IsMatch(phaseCRightValue, @"^-?\d+$")
              ? phaseCRightValue
              : Math.Round(Convert.ToDouble(phaseCRightValue)).ToString();
        }

        phase_a_left_tag.Add(phaseALeftTag);
        phase_b_left_tag.Add(phaseBLeftTag);
        phase_a_right_tag.Add(phaseARightTag);
        phase_b_right_tag.Add(phaseBRightTag);
        phase_c_left_tag.Add(phaseCLeftTag);
        phase_c_right_tag.Add(phaseCRightTag);

        description_left_tags.Add(descriptionLeftTag);
        description_right_tags.Add(descriptionRightTag);

        // Checks for Left Side
        bool hasCommaInPhaseLeft =
          phaseALeftValue.Contains(";")
          || phaseBLeftValue.Contains(";")
          || phaseCLeftValue.Contains(";");
        bool shouldDuplicateLeft = hasCommaInPhaseLeft;

        // Checks for Right Side
        bool hasCommaInPhaseRight =
          phaseARightValue.Contains(";")
          || phaseBRightValue.Contains(";")
          || phaseCRightValue.Contains(";");
        bool shouldDuplicateRight = hasCommaInPhaseRight;

        // Handling Phase A Left
        if (phaseALeftValue.Contains(";")) {
          var splitValues = phaseALeftValue.Split(';').Select(str => str.Trim()).ToArray();
          phase_a_left.AddRange(splitValues);
        }
        else {
          phase_a_left.Add(phaseALeftValue);
          phase_a_left.Add("0"); // Default value
        }

        // Handling Phase B Left
        if (phaseBLeftValue.Contains(";")) {
          var splitValues = phaseBLeftValue.Split(';').Select(str => str.Trim()).ToArray();
          phase_b_left.AddRange(splitValues);
        }
        else {
          phase_b_left.Add(phaseBLeftValue);
          phase_b_left.Add("0"); // Default value
        }

        // Handling Phase A Right
        if (phaseARightValue.Contains(";")) {
          var splitValues = phaseARightValue.Split(';').Select(str => str.Trim()).ToArray();
          phase_a_right.AddRange(splitValues);
        }
        else {
          phase_a_right.Add(phaseARightValue);
          phase_a_right.Add("0"); // Default value
        }

        // Handling Phase B Right
        if (phaseBRightValue.Contains(";")) {
          var splitValues = phaseBRightValue.Split(';').Select(str => str.Trim()).ToArray();
          phase_b_right.AddRange(splitValues);
        }
        else {
          phase_b_right.Add(phaseBRightValue);
          phase_b_right.Add("0"); // Default value
        }

        if (PHASE_SUM_GRID.Columns.Count > 2) {
          // Handling Phase C Left
          if (phaseCLeftValue.Contains(";")) {
            var splitValues = phaseCLeftValue.Split(';').Select(str => str.Trim()).ToArray();
            phase_c_left.AddRange(splitValues);
          }
          else {
            phase_c_left.Add(phaseCLeftValue);
            phase_c_left.Add("0"); // Default value
          }

          // Handling Phase C Right
          if (phaseCRightValue.Contains(";")) {
            var splitValues = phaseCRightValue.Split(';').Select(str => str.Trim()).ToArray();
            phase_c_right.AddRange(splitValues);
          }
          else {
            phase_c_right.Add(phaseCRightValue);
            phase_c_right.Add("0"); // Default value
          }
        }

        if (descriptionLeftValue.Contains(";")) {
          // If it contains a comma, split and add both values
          var splitValues = descriptionLeftValue.Split(';').Select(str => str.Trim()).ToArray();
          description_left.AddRange(splitValues);
          circuit_left.Add(circuitLeftValue + "A");
          circuit_left.Add(circuitLeftValue + "B");
        }
        else {
          description_left.Add(descriptionLeftValue);
          description_left.Add(shouldDuplicateLeft ? descriptionLeftValue : "SPACE");

          if (shouldDuplicateLeft) {
            circuit_left.Add(circuitLeftValue + "A");
            circuit_left.Add(circuitLeftValue + "B");
          }
          else {
            circuit_left.Add(circuitLeftValue);
            circuit_left.Add("");
          }
        }

        if (breakerLeftValue.Contains(";")) {
          // If it contains a comma, split and add both values
          var splitValues = breakerLeftValue.Split(';').Select(str => str.Trim()).ToArray();
          breaker_left.AddRange(splitValues);
        }
        else {
          breaker_left.Add(breakerLeftValue);
          breaker_left.Add(shouldDuplicateLeft ? breakerLeftValue : "");
        }

        if (descriptionRightValue.Contains(";")) {
          // If it contains a comma, split and add both values
          var splitValues = descriptionRightValue.Split(';').Select(str => str.Trim()).ToArray();
          description_right.AddRange(splitValues);
          circuit_right.Add(circuitRightValue + "A");
          circuit_right.Add(circuitRightValue + "B");
        }
        else {
          description_right.Add(descriptionRightValue);
          description_right.Add(shouldDuplicateRight ? descriptionRightValue : "SPACE");

          if (shouldDuplicateRight) {
            circuit_right.Add(circuitRightValue + "A");
            circuit_right.Add(circuitRightValue + "B");
          }
          else {
            circuit_right.Add(circuitRightValue);
            circuit_right.Add("");
          }
        }

        if (breakerRightValue.Contains(";")) {
          // If it contains a comma, split and add both values
          var splitValues = breakerRightValue.Split(';').Select(str => str.Trim()).ToArray();
          breaker_right.AddRange(splitValues);
        }
        else {
          breaker_right.Add(breakerRightValue);
          breaker_right.Add(shouldDuplicateRight ? breakerRightValue : "");
        }

        // Left Side
        description_left_highlights.Add(false);
        breaker_left_highlights.Add(false);

        // Right Side
        description_right_highlights.Add(false);
        breaker_right_highlights.Add(false);

        // Default Values for Left Side
        description_left_highlights.Add(false);
        breaker_left_highlights.Add(false);

        // Default Values for Right Side
        description_right_highlights.Add(false);
        breaker_right_highlights.Add(false);
      }

      panel.Add("description_left_highlights", description_left_highlights);
      panel.Add("description_right_highlights", description_right_highlights);
      panel.Add("breaker_left_highlights", breaker_left_highlights);
      panel.Add("breaker_right_highlights", breaker_right_highlights);
      panel.Add("description_left", description_left);
      panel.Add("description_right", description_right);
      panel.Add("phase_a_left", phase_a_left);
      panel.Add("phase_b_left", phase_b_left);
      panel.Add("phase_a_right", phase_a_right);
      panel.Add("phase_b_right", phase_b_right);

      if (PHASE_SUM_GRID.Columns.Count > 2) {
        panel.Add("phase_c_left", phase_c_left);
        panel.Add("phase_c_right", phase_c_right);
      }

      panel.Add("breaker_left", breaker_left);
      panel.Add("breaker_right", breaker_right);
      panel.Add("circuit_left", circuit_left);
      panel.Add("circuit_right", circuit_right);
      panel.Add("phase_a_left_tag", phase_a_left_tag);
      panel.Add("phase_b_left_tag", phase_b_left_tag);
      panel.Add("phase_a_right_tag", phase_a_right_tag);
      panel.Add("phase_b_right_tag", phase_b_right_tag);

      if (PHASE_SUM_GRID.Columns.Count > 2) {
        panel.Add("phase_c_left_tag", phase_c_left_tag);
        panel.Add("phase_c_right_tag", phase_c_right_tag);
      }

      panel.Add("description_left_tags", description_left_tags);
      panel.Add("description_right_tags", description_right_tags);

      panel.Add("notes", notesStorage);

      return panel;
    }

    public double CalculateTotalVA(double sum) {
      return Math.Round(sum, 0);
    }

    public double CalculatePanelLoad(double sum) {
      return Math.Round(sum / 1000, 1);
    }

    public double GetPanelLoad() {
      if (PANEL_LOAD_GRID != null && PANEL_LOAD_GRID.Rows.Count > 0 && PANEL_LOAD_GRID.Columns.Count > 0) {
        object cellValue = PANEL_LOAD_GRID.Rows[0].Cells[0].Value;
        if (cellValue != null && double.TryParse(cellValue.ToString(), out double totalKVA)) {
          return totalKVA;
        }
      }
      return 0.0;
    }

    public List<string> GetSubPanels() {
      List<string> subPanels = new List<string>();
      string pattern = @"(PANEL|SUBPANEL)\s+(?:'?([^']+)'?)";
      Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);

      foreach (DataGridViewRow row in PANEL_GRID.Rows) {
        foreach (DataGridViewCell cell in row.Cells) {
          string cellValue = cell.Value?.ToString() ?? "";
          MatchCollection matches = regex.Matches(cellValue);

          foreach (Match match in matches) {
            if (match.Groups.Count > 2) {
              string panelName = match.Groups[2].Value;
              if (!subPanels.Contains(panelName)) {
                subPanels.Add(panelName.ToUpper());

              }
            }
          }
        }
      }

      return subPanels;
    }

    public string GetPanelName() {
      return PANEL_NAME_INPUT.Text;
    }

    public double StoreItemsAndWattage(string note) {
      int phaseCount = PHASE_SUM_GRID.ColumnCount;
      string[] columnNames = GetColumnNames(phaseCount);

      List<PanelItem> items = new List<PanelItem>();

      // Process left side
      ProcessSide(columnNames, "left", items, note);

      // Process right side
      ProcessSide(columnNames, "right", items, note);

      return items.Count > 0
            ? items.Max(item => item.Wattage)
            : 0;
    }

    private void ProcessSide(string[] columnNames, string side, List<PanelItem> items, string note) {
      int startIndex = side == "left" ? 0 : 1;
      for (int rowIndex = 0; rowIndex < PANEL_GRID.Rows.Count; rowIndex++) {
        DataGridViewRow row = PANEL_GRID.Rows[rowIndex];
        for (int i = startIndex; i < columnNames.Length; i += 2) {
          string colName = columnNames[i];
          bool descriptionHasNote = DescriptionHasNote(row, side, note);
          int breakerValue = BreakerToInt(row, side);

          if (row.Cells[colName].Value != null && descriptionHasNote && breakerValue > 3) {
            double phaseValue;
            if (!TryParseDouble(row.Cells[colName].Value, out phaseValue)) {
              continue;
            }
            string description = GetDescription(row, side);
            PanelItem item = new PanelItem { Description = description };

            // Check for condition 1
            if (rowIndex + 2 < PANEL_GRID.Rows.Count) {
              DataGridViewRow secondRow = PANEL_GRID.Rows[rowIndex + 1];
              DataGridViewRow thirdRow = PANEL_GRID.Rows[rowIndex + 2];
              int secondBreakerValue = BreakerToInt(secondRow, side);
              string thirdBreakerValue = thirdRow.Cells[$"breaker_{side}"].Value?.ToString();

              if (secondBreakerValue == 0 && thirdBreakerValue == "3") {
                item.Wattage = phaseValue * 3 / 1.732;
                item.Poles = 3;
                rowIndex += 2;
                items.Add(item);
                continue;
              }
            }

            // Check for condition 2
            if (rowIndex + 1 < PANEL_GRID.Rows.Count) {
              DataGridViewRow secondRow = PANEL_GRID.Rows[rowIndex + 1];
              string secondBreakerValue = secondRow.Cells[$"breaker_{side}"].Value?.ToString();

              if (secondBreakerValue == "2") {
                item.Wattage = phaseValue * 2;
                item.Poles = 2;
                rowIndex += 1; // Skip 1 row
                items.Add(item);
                continue;
              }
            }

            // Condition 3 (default case)
            item.Wattage = phaseValue;
            item.Poles = 1;
            items.Add(item);
          }
        }
      }
    }

    private bool TryParseDouble(object value, out double result) {
      result = 0;
      if (value == null) return false;

      string stringValue = value.ToString().Trim();
      if (string.IsNullOrEmpty(stringValue)) return false;

      // Try parsing with invariant culture (uses period as decimal separator)
      if (double.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
        return true;

      // If that fails, try parsing with the current culture
      return double.TryParse(stringValue, NumberStyles.Any, CultureInfo.CurrentCulture, out result);
    }

    private bool DescriptionHasNote(DataGridViewRow row, string side, string note) {
      string columnName = $"description_{side}";
      DataGridViewCell cell = row.Cells[columnName];

      if (cell.Tag == null) {
        return false;
      }

      string tagValue = cell.Tag.ToString();
      return tagValue.Contains(note);
    }

    private int BreakerToInt(DataGridViewRow row, string side) {
      string columnName = $"breaker_{side}";
      var cellValue = row.Cells[columnName].Value;

      if (cellValue == null || string.IsNullOrEmpty(cellValue.ToString())) {
        return 0;
      }

      string breakerValue = cellValue.ToString();

      if (breakerValue.Contains(",")) {
        breakerValue = breakerValue.Split(',')[0];
      }

      if (int.TryParse(breakerValue, out int result)) {
        return result;
      }

      return 0;
    }

    private string GetDescription(DataGridViewRow row, string side) {
      return row.Cells[$"description_{side}"].Value?.ToString() ?? string.Empty;
    }

    public double CalculateWattageSum(string note) {
      int phaseCount = PHASE_SUM_GRID.ColumnCount;
      if (phaseCount < 2 || phaseCount > 3) {
        throw new ArgumentException("Unsupported phase count. Must be 2 or 3.");
      }
      return CalculateWattageSumForPhases(phaseCount, note);
    }

    private double CalculateWattageSumForPhases(int phaseCount, string note) {
      string[] columnNames = GetColumnNames(phaseCount);
      double[] sums = new double[phaseCount];

      foreach (DataGridViewRow row in PANEL_GRID.Rows) {
        for (int i = 0; i < columnNames.Length; i += 2) {
          for (int j = 0; j < 2; j++) {
            string colName = columnNames[i + j];
            if (row.Cells[colName].Value != null) {
              bool hasNoteApplied = BreakerContainsNote(row.Index, colName, note);
              if (hasNoteApplied) {
                sums[i / 2] += ParseAndSumCell(row.Cells[colName].Value.ToString(), 1);
              }
            }
          }
        }
      }

      return sums.Sum();
    }

    public void CalculateBreakerLoad() {
      int phaseCount = PHASE_SUM_GRID.ColumnCount;
      if (phaseCount < 2 || phaseCount > 3) {
        throw new ArgumentException("Unsupported phase count. Must be 2 or 3.");
      }

      CalculateBreakerLoadForPhases(phaseCount);
    }

    private bool RowIsSinglePhase(int i, string side) {
      if (i == PANEL_GRID.Rows.Count - 1) {
        if (String.IsNullOrEmpty((PANEL_GRID.Rows[i].Cells[$"breaker_{side}"].Value as string))) {
          return false;
        }
        if (PANEL_GRID.Rows[i].Cells[$"breaker_{side}"].Value as string == "3") {
          return false;
        }
        if (PANEL_GRID.Rows[i].Cells[$"breaker_{side}"].Value as string == "2") {
          return false;
        }
        return true;
      }
      else if (i == PANEL_GRID.Rows.Count - 2) {
        if (String.IsNullOrEmpty((PANEL_GRID.Rows[i].Cells[$"breaker_{side}"].Value as string))) {
          return false;
        }
        if (PANEL_GRID.Rows[i].Cells[$"breaker_{side}"].Value as string == "3") {
          return false;
        }
        if (PANEL_GRID.Rows[i].Cells[$"breaker_{side}"].Value as string == "2") {
          return false;
        }
        if (PANEL_GRID.Rows[i + 1].Cells[$"breaker_{side}"].Value as string == "3") {
          return false;
        }
        if (PANEL_GRID.Rows[i + 1].Cells[$"breaker_{side}"].Value as string == "2") {
          return false;
        }
        return true;
      }
      else {
        if (String.IsNullOrEmpty((PANEL_GRID.Rows[i].Cells[$"breaker_{side}"].Value as string))) {
          return false;
        }
        if (PANEL_GRID.Rows[i].Cells[$"breaker_{side}"].Value as string == "3") {
          return false;
        }
        if (PANEL_GRID.Rows[i].Cells[$"breaker_{side}"].Value as string == "2") {
          return false;
        }
        if (PANEL_GRID.Rows[i + 1].Cells[$"breaker_{side}"].Value as string == "3") {
          return false;
        }
        if (PANEL_GRID.Rows[i + 1].Cells[$"breaker_{side}"].Value as string == "2") {
          return false;
        }
        if (PANEL_GRID.Rows[i + 2].Cells[$"breaker_{side}"].Value as string == "3") {
          return false;
        }
        return true;
      }
    }

    public void SetWarnings() {
      PHASE_COMBOBOX.Enabled = true;
      WIRE_COMBOBOX.Enabled = true;
      this.contains3PhEquip = false;
      PHASE_WARNING_LABEL.Visible = false;
      HIGH_LEG_WARNING_LEFT_LABEL.Visible = false;
      HIGH_LEG_WARNING_RIGHT_LABEL.Visible = false;
      for (int i = 0; i < PANEL_GRID.Rows.Count; i++) {
        DataGridViewRow row = PANEL_GRID.Rows[i];
        if (!String.IsNullOrEmpty(row.Cells["breaker_left"].Value as string) && row.Cells["breaker_left"].Value as string == "3" && is3Ph) {
          PHASE_COMBOBOX.Enabled = false;
          WIRE_COMBOBOX.Enabled = false;
          this.contains3PhEquip = true;
          PHASE_WARNING_LABEL.Visible = true;
        }
        else if (!String.IsNullOrEmpty(row.Cells["breaker_right"].Value as string) && row.Cells["breaker_right"].Value as string == "3" && is3Ph) {
          PHASE_COMBOBOX.Enabled = false;
          WIRE_COMBOBOX.Enabled = false;
          this.contains3PhEquip = true;
          PHASE_WARNING_LABEL.Visible = true;
        }
        else if (!String.IsNullOrEmpty(FED_FROM_TEXTBOX.Text)) {
          PHASE_COMBOBOX.Enabled = false;
          WIRE_COMBOBOX.Enabled = false;
          PHASE_WARNING_LABEL.Visible = true;
        }
        if (is3Ph && i % 3 == 2
          && LINE_VOLTAGE_COMBOBOX.Text == "240"
          && RowIsSinglePhase(i, "left")) {
          HIGH_LEG_WARNING_LEFT_LABEL.Visible = true;
          row.Cells["breaker_left"].Style.BackColor = Color.Crimson;
          row.Cells["breaker_left"].Style.ForeColor = Color.White;
        }
        else {
          row.Cells["breaker_left"].Style.BackColor = Color.White;
          row.Cells["breaker_left"].Style.ForeColor = Color.Black;
        }
        if (is3Ph && i % 3 == 2
          && LINE_VOLTAGE_COMBOBOX.Text == "240"
          && RowIsSinglePhase(i, "right")) {
          HIGH_LEG_WARNING_RIGHT_LABEL.Visible = true;
          row.Cells["breaker_right"].Style.BackColor = Color.Crimson;
          row.Cells["breaker_right"].Style.ForeColor = Color.White;
        }
        else {
          row.Cells["breaker_right"].Style.BackColor = Color.White;
          row.Cells["breaker_right"].Style.ForeColor = Color.Black;
        }
      }
      
    }

    private void CalculateBreakerLoadForPhases(int phaseCount) {
      string[] columnNames = GetColumnNames(phaseCount);
      int breakersWithKitchenDemand = BreakersWithNote("KITCHEN DEMAND");
      double demandFactor = KitchenDemandFactor(breakersWithKitchenDemand);
      double[] sums = new double[phaseCount];
      foreach (DataGridViewRow row in PANEL_GRID.Rows)  {
        for (int i = 0; i < columnNames.Length; i += 2) {
          for (int j = 0; j < 2; j++) {
            string colName = columnNames[i + j];
            if (row.Cells[colName].Value != null) {
              bool hasKitchenDemandApplied = BreakerContainsNote(row.Index, colName, "KITCHEN DEMAND");
              sums[i / 2] += ParseAndSumCell(
                  row.Cells[colName].Value.ToString(),
                  hasKitchenDemandApplied ? demandFactor : 1.0
              );
            }
          }
        }
      }
      SetWarnings();
      for (int i = 0; i < phaseCount; i++) {
        if (PHASE_SUM_GRID.Rows.Count > 0) {
          PHASE_SUM_GRID.Rows[0].Cells[i].Value = sums[i];
        }
      }
    }

    private string[] GetColumnNames(int phaseCount) {
      if (phaseCount == 2) {
        return new[] { "phase_a_left", "phase_a_right", "phase_b_left", "phase_b_right" };
      }
      else {
        return new[] { "phase_a_left", "phase_a_right", "phase_b_left", "phase_b_right", "phase_c_left", "phase_c_right" };
      }
    }

    private int BreakersWithNote(string note) {
      return PANEL_GRID.Rows.Cast<DataGridViewRow>()
          .Sum(row => new[] { "description_left", "description_right" }
              .Count(colName => CellHasNote(colName, row, note)));
    }

    private bool CellHasNote(string columnName, DataGridViewRow row, string note) {
      if (row == null || !PANEL_GRID.Columns.Contains(columnName))
        return false;

      var cell = row.Cells[columnName];
      if (cell == null)
        return false;

      string cellValueString = cell.Value?.ToString() ?? "";
      string cellTagString = cell.Tag?.ToString() ?? "";

      return !string.IsNullOrEmpty(cellValueString) && cellTagString.Contains(note);
    }

    private double ParseAndSumCell(
      string cellValue,
      double demandFactor
    ) {
      double sum = 0;
      if (!string.IsNullOrEmpty(cellValue)) {
        var parts = cellValue.Split(';');
        foreach (var part in parts) {
          if (double.TryParse(part, out double value)) {
            if (demandFactor != 1.00) {
              sum += value * demandFactor;
            }
            else {
              sum += value;
            }
          }
        }
      }
      return Math.Ceiling(sum);
    }

    private double KitchenDemandFactor(int numberOfBreakersWithKitchenDemand) {
      if (numberOfBreakersWithKitchenDemand == 1 || numberOfBreakersWithKitchenDemand == 2) {
        return 1.00;
      }
      else if (numberOfBreakersWithKitchenDemand == 3) {
        return 0.90;
      }
      else if (numberOfBreakersWithKitchenDemand == 4) {
        return 0.80;
      }
      else if (numberOfBreakersWithKitchenDemand == 5) {
        return 0.70;
      }
      else if (numberOfBreakersWithKitchenDemand >= 6) {
        return 0.65;
      }
      else {
        return 1.00;
      }
    }

    private void ListenFor3pRowsAdded(DataGridViewRowsAddedEventArgs e) {
      Color grayColor = Color.LightGray;

      for (int i = 0; i < e.RowCount; i++) {
        int rowIndex = e.RowIndex + i;

        // Set common column values
        PANEL_GRID.Rows[rowIndex].Cells["description_left"].Value = "SPARE";
        PANEL_GRID.Rows[rowIndex].Cells["breaker_left"].Value = "20";
        PANEL_GRID.Rows[rowIndex].Cells["circuit_left"].Value = ((rowIndex + 1) * 2) - 1;
        PANEL_GRID.Rows[rowIndex].Cells["circuit_right"].Value = (rowIndex + 1) * 2;
        PANEL_GRID.Rows[rowIndex].Cells["breaker_right"].Value = "20";
        PANEL_GRID.Rows[rowIndex].Cells["description_right"].Value = "SPARE";

        // Determine the row pattern (zig-zag) for gray background
        int pattern = rowIndex % 3;

        // Apply pattern for two sets of columns based on the row pattern
        if (pattern == 0) {
          PANEL_GRID.Rows[rowIndex].Cells["phase_a_left"].Style.BackColor = grayColor;
          PANEL_GRID.Rows[rowIndex].Cells["phase_a_right"].Style.BackColor = grayColor;
        }
        else if (pattern == 1) {
          PANEL_GRID.Rows[rowIndex].Cells["phase_b_left"].Style.BackColor = grayColor;
          PANEL_GRID.Rows[rowIndex].Cells["phase_b_right"].Style.BackColor = grayColor;
        }
        else {
          PANEL_GRID.Rows[rowIndex].Cells["phase_c_left"].Style.BackColor = grayColor;
          PANEL_GRID.Rows[rowIndex].Cells["phase_c_right"].Style.BackColor = grayColor;
        }
      }
    }

    private void ListenFor2pRowsAdded(DataGridViewRowsAddedEventArgs e) {
      for (int i = 0; i < e.RowCount; i++) {
        int rowIndex = e.RowIndex + i;

        // Set common column values
        PANEL_GRID.Rows[rowIndex].Cells["description_left"].Value = "SPARE";
        PANEL_GRID.Rows[rowIndex].Cells["breaker_left"].Value = "20";
        PANEL_GRID.Rows[rowIndex].Cells["circuit_left"].Value = ((rowIndex + 1) * 2) - 1;
        PANEL_GRID.Rows[rowIndex].Cells["circuit_right"].Value = (rowIndex + 1) * 2;
        PANEL_GRID.Rows[rowIndex].Cells["breaker_right"].Value = "20";
        PANEL_GRID.Rows[rowIndex].Cells["description_right"].Value = "SPARE";

        // Zig-zag pattern for columns 2, 3, 8, and 9
        if ((rowIndex + 1) % 2 == 1) // Odd rows
        {
          PANEL_GRID.Rows[rowIndex].Cells["phase_a_left"].Style.BackColor = Color.LightGray; // Column 2
          PANEL_GRID.Rows[rowIndex].Cells["phase_a_right"].Style.BackColor = Color.LightGray; // Column 8
        }
        else // Even rows
        {
          PANEL_GRID.Rows[rowIndex].Cells["phase_b_left"].Style.BackColor = Color.LightGray; // Column 3
          PANEL_GRID.Rows[rowIndex].Cells["phase_b_right"].Style.BackColor = Color.LightGray; // Column 9
        }
      }
    }

    private void Color3pPanel(object sender, EventArgs e) {
      Color backColor1 = Color.White;
      Color backColor2 = Color.AliceBlue;
      Color foreColor1 = Color.Black;
      Color foreColor2 = Color.Black;
      Color blockColor = Color.SlateGray;
      Color phaseColor = Color.LightGray;
      Color hiLegColor = Color.LightBlue;
      double lineVoltage = SafeConvertToDouble(LINE_VOLTAGE_COMBOBOX.Text);
      if (!DISTRIBUTION_SECTION_CHECKBOX.Checked) {
        backColor2 = Color.White;
      }
      for (int i = 0; i < PANEL_GRID.Rows.Count; i++) {

        if (DISTRIBUTION_SECTION_CHECKBOX.Checked) {
          if (i % 6 >= 3) {
            PANEL_GRID.Rows[i].Cells["description_left"].Style.BackColor = backColor2;
            PANEL_GRID.Rows[i].Cells["description_right"].Style.BackColor = backColor2;
            PANEL_GRID.Rows[i].Cells["breaker_left"].Style.BackColor = backColor2;
            PANEL_GRID.Rows[i].Cells["breaker_right"].Style.BackColor = backColor2;
            if (i % 3 == 2) {
              PANEL_GRID.Rows[i].Cells["phase_a_left"].Style.BackColor = backColor2;
              PANEL_GRID.Rows[i].Cells["phase_b_left"].Style.BackColor = backColor2;
              PANEL_GRID.Rows[i].Cells["phase_a_right"].Style.BackColor = backColor2;
              PANEL_GRID.Rows[i].Cells["phase_b_right"].Style.BackColor = backColor2;
            }
            else if (i % 3 == 1) {
              PANEL_GRID.Rows[i].Cells["phase_a_left"].Style.BackColor = backColor2;
              PANEL_GRID.Rows[i].Cells["phase_c_left"].Style.BackColor = backColor2;
              PANEL_GRID.Rows[i].Cells["phase_a_right"].Style.BackColor = backColor2;
              PANEL_GRID.Rows[i].Cells["phase_c_right"].Style.BackColor = backColor2;
            }
            else if (i % 3 == 0) {
              PANEL_GRID.Rows[i].Cells["phase_b_left"].Style.BackColor = backColor2;
              PANEL_GRID.Rows[i].Cells["phase_c_left"].Style.BackColor = backColor2;
              PANEL_GRID.Rows[i].Cells["phase_b_right"].Style.BackColor = backColor2;
              PANEL_GRID.Rows[i].Cells["phase_c_right"].Style.BackColor = backColor2;
            }
          }
          else {
            PANEL_GRID.Rows[i].Cells["breaker_left"].Style.BackColor = backColor1;
            PANEL_GRID.Rows[i].Cells["breaker_right"].Style.BackColor = backColor1;
          }
          if (i % 3 != 0) {
            if (i % 6 >= 3) {
              PANEL_GRID.Rows[i].Cells["description_left"].Style.BackColor = backColor2;
              PANEL_GRID.Rows[i].Cells["description_left"].Style.ForeColor = backColor2;
              PANEL_GRID.Rows[i].Cells["description_right"].Style.BackColor = backColor2;
              PANEL_GRID.Rows[i].Cells["description_right"].Style.ForeColor = backColor2;
              PANEL_GRID.Rows[i].Cells["description_left"].ReadOnly = true;
              PANEL_GRID.Rows[i].Cells["description_right"].ReadOnly = true;
            }
            else {
              PANEL_GRID.Rows[i].Cells["description_left"].Style.BackColor = backColor1;
              PANEL_GRID.Rows[i].Cells["description_left"].Style.ForeColor = backColor1;
              PANEL_GRID.Rows[i].Cells["description_right"].Style.BackColor = backColor1;
              PANEL_GRID.Rows[i].Cells["description_right"].Style.ForeColor = backColor1;
              PANEL_GRID.Rows[i].Cells["description_left"].ReadOnly = true;
              PANEL_GRID.Rows[i].Cells["description_right"].ReadOnly = true;
            }
          }
          PANEL_GRID.Rows[i].Cells["circuit_left"].Style.BackColor = blockColor;
          PANEL_GRID.Rows[i].Cells["circuit_right"].Style.BackColor = blockColor;
          PANEL_GRID.Rows[i].Cells["circuit_left"].Style.ForeColor = blockColor;
          PANEL_GRID.Rows[i].Cells["circuit_right"].Style.ForeColor = blockColor;
          PANEL_GRID.Rows[i].Cells["circuit_left"].ReadOnly = true;
          PANEL_GRID.Rows[i].Cells["circuit_right"].ReadOnly = true;
          PANEL_NAME_LABEL.Text = "DISTRIBUTION SECTION";
          PANEL_NAME_LABEL.Location = new Point(57, 74);
          this.mainForm.PANEL_NAME_INPUT_TextChanged(sender, e, PANEL_NAME_INPUT.Text, true);
          CREATE_PANEL_BUTTON.Visible = false;
          CREATE_LOAD_SUMMARY_BUTTON.Visible = true;
          ADD_ALL_PANELS_BUTTON.Visible = true;
          if (GetPanelLoad() > 0) {
            ADD_ALL_PANELS_BUTTON.Enabled = false;
          }
          else {
            ADD_ALL_PANELS_BUTTON.Enabled = true;
          }
          SAFETY_FACTOR_CHECKBOX.Enabled = true;
        }
        else {
          if (i % 3 == 0) { // phase a shaded
            PANEL_GRID.Rows[i].Cells["phase_a_left"].Style.BackColor = phaseColor;
            PANEL_GRID.Rows[i].Cells["phase_a_right"].Style.BackColor = phaseColor;
            PANEL_GRID.Rows[i].Cells["phase_b_left"].Style.BackColor = backColor1;
            PANEL_GRID.Rows[i].Cells["phase_b_right"].Style.BackColor = backColor1;
            PANEL_GRID.Rows[i].Cells["phase_c_left"].Style.BackColor = backColor1;
            PANEL_GRID.Rows[i].Cells["phase_c_right"].Style.BackColor = backColor1;
          }
          else if (i % 3 == 1) { // phase b shaded
            PANEL_GRID.Rows[i].Cells["phase_a_left"].Style.BackColor = backColor1;
            PANEL_GRID.Rows[i].Cells["phase_a_right"].Style.BackColor = backColor1;
            PANEL_GRID.Rows[i].Cells["phase_b_left"].Style.BackColor = phaseColor;
            PANEL_GRID.Rows[i].Cells["phase_b_right"].Style.BackColor = phaseColor;
            PANEL_GRID.Rows[i].Cells["phase_c_left"].Style.BackColor = backColor1;
            PANEL_GRID.Rows[i].Cells["phase_c_right"].Style.BackColor = backColor1;
          }
          else { // phase c shaded
            PANEL_GRID.Rows[i].Cells["phase_a_left"].Style.BackColor = backColor1;
            PANEL_GRID.Rows[i].Cells["phase_a_right"].Style.BackColor = backColor1;
            PANEL_GRID.Rows[i].Cells["phase_b_left"].Style.BackColor = backColor1;
            PANEL_GRID.Rows[i].Cells["phase_b_right"].Style.BackColor = backColor1;
            if (lineVoltage != 240) {
              PANEL_GRID.Rows[i].Cells["phase_c_left"].Style.BackColor = phaseColor;
              PANEL_GRID.Rows[i].Cells["phase_c_right"].Style.BackColor = phaseColor;
              PANEL_GRID.Rows[i].Cells["phase_c_left"].Style.ForeColor = foreColor1;
              PANEL_GRID.Rows[i].Cells["phase_c_right"].Style.ForeColor = foreColor1;
            }
            else {
              PANEL_GRID.Rows[i].Cells["phase_c_left"].Style.BackColor = hiLegColor;
              PANEL_GRID.Rows[i].Cells["phase_c_right"].Style.BackColor = hiLegColor;
              PANEL_GRID.Rows[i].Cells["phase_c_left"].Style.ForeColor = foreColor2;
              PANEL_GRID.Rows[i].Cells["phase_c_right"].Style.ForeColor = foreColor2;
            }
          }
          PANEL_GRID.Rows[i].Cells["description_left"].Style.BackColor = backColor1;
          PANEL_GRID.Rows[i].Cells["description_right"].Style.BackColor = backColor1;
          PANEL_GRID.Rows[i].Cells["description_left"].Style.ForeColor = foreColor1;
          PANEL_GRID.Rows[i].Cells["description_right"].Style.ForeColor = foreColor1;
          PANEL_GRID.Rows[i].Cells["circuit_left"].Style.BackColor = backColor1;
          PANEL_GRID.Rows[i].Cells["circuit_right"].Style.BackColor = backColor1;
          PANEL_GRID.Rows[i].Cells["circuit_left"].Style.ForeColor = foreColor1;
          PANEL_GRID.Rows[i].Cells["circuit_right"].Style.ForeColor = foreColor1;
          PANEL_GRID.Rows[i].Cells["circuit_left"].ReadOnly = false;
          PANEL_GRID.Rows[i].Cells["circuit_right"].ReadOnly = false;
          PANEL_GRID.Rows[i].Cells["description_left"].ReadOnly = false;
          PANEL_GRID.Rows[i].Cells["description_right"].ReadOnly = false;
          PANEL_GRID.Rows[i].Cells["breaker_left"].Style.BackColor = backColor1;
          PANEL_GRID.Rows[i].Cells["breaker_right"].Style.BackColor = backColor1;
          PANEL_NAME_LABEL.Text = "PANEL";
          PANEL_NAME_LABEL.Location = new Point(150, 74);
          this.mainForm.PANEL_NAME_INPUT_TextChanged(sender, e, PANEL_NAME_INPUT.Text, false);
          CREATE_PANEL_BUTTON.Visible = true;
          CREATE_LOAD_SUMMARY_BUTTON.Visible = false;
          ADD_ALL_PANELS_BUTTON.Visible = false;
          SAFETY_FACTOR_CHECKBOX.Enabled = false;
        }
      }
    }

    private void Color2pPanel(object sender, EventArgs e) {
      Color backColor1 = Color.White;
      Color backColor2 = Color.AliceBlue;
      Color foreColor1 = Color.Black;
      Color foreColor2 = Color.Black;
      Color blockColor = Color.SlateGray;
      Color phaseColor = Color.LightGray;
      if (!DISTRIBUTION_SECTION_CHECKBOX.Checked) {
        backColor2 = Color.White;
      }
      for (int i = 0; i < PANEL_GRID.Rows.Count; i++) {
        if (DISTRIBUTION_SECTION_CHECKBOX.Checked) {
          if (i % 4 >= 2) {
            PANEL_GRID.Rows[i].Cells["description_left"].Style.BackColor = backColor2;
            PANEL_GRID.Rows[i].Cells["description_right"].Style.BackColor = backColor2;
            PANEL_GRID.Rows[i].Cells["description_left"].Style.ForeColor = foreColor2;
            PANEL_GRID.Rows[i].Cells["description_right"].Style.ForeColor = foreColor2;
            PANEL_GRID.Rows[i].Cells["breaker_left"].Style.BackColor = backColor2;
            PANEL_GRID.Rows[i].Cells["breaker_right"].Style.BackColor = backColor2;
            if (i % 2 == 1) {
              PANEL_GRID.Rows[i].Cells["phase_a_left"].Style.BackColor = backColor2;
              PANEL_GRID.Rows[i].Cells["phase_a_right"].Style.BackColor = backColor2;
            }
            else if (i % 2 == 0) {
              PANEL_GRID.Rows[i].Cells["phase_b_left"].Style.BackColor = backColor2;
              PANEL_GRID.Rows[i].Cells["phase_b_right"].Style.BackColor = backColor2;
            }
          }
          else {
            PANEL_GRID.Rows[i].Cells["breaker_left"].Style.BackColor = backColor1;
            PANEL_GRID.Rows[i].Cells["breaker_right"].Style.BackColor = backColor1;
          }
          if (i % 2 != 0) {
            if (i % 4 >= 2) {
              PANEL_GRID.Rows[i].Cells["description_left"].Style.BackColor = backColor2;
              PANEL_GRID.Rows[i].Cells["description_left"].Style.ForeColor = backColor2;
              PANEL_GRID.Rows[i].Cells["description_right"].Style.BackColor = backColor2;
              PANEL_GRID.Rows[i].Cells["description_right"].Style.ForeColor = backColor2;
              PANEL_GRID.Rows[i].Cells["description_left"].ReadOnly = true;
              PANEL_GRID.Rows[i].Cells["description_right"].ReadOnly = true;
            }
            else {
              PANEL_GRID.Rows[i].Cells["description_left"].Style.BackColor = backColor1;
              PANEL_GRID.Rows[i].Cells["description_left"].Style.ForeColor = backColor1;
              PANEL_GRID.Rows[i].Cells["description_right"].Style.BackColor = backColor1;
              PANEL_GRID.Rows[i].Cells["description_right"].Style.ForeColor = backColor1;
              PANEL_GRID.Rows[i].Cells["description_left"].ReadOnly = true;
              PANEL_GRID.Rows[i].Cells["description_right"].ReadOnly = true;
            }
          }
          PANEL_GRID.Rows[i].Cells["circuit_left"].Style.BackColor = blockColor;
          PANEL_GRID.Rows[i].Cells["circuit_right"].Style.BackColor = blockColor;
          PANEL_GRID.Rows[i].Cells["circuit_left"].Style.ForeColor = blockColor;
          PANEL_GRID.Rows[i].Cells["circuit_right"].Style.ForeColor = blockColor;
          PANEL_GRID.Rows[i].Cells["circuit_left"].ReadOnly = true;
          PANEL_GRID.Rows[i].Cells["circuit_right"].ReadOnly = true;
          PANEL_NAME_LABEL.Text = "DISTRIBUTION SECTION";
          PANEL_NAME_LABEL.Location = new Point(57, 74);
          this.mainForm.PANEL_NAME_INPUT_TextChanged(sender, e, PANEL_NAME_INPUT.Text, true);
          CREATE_PANEL_BUTTON.Visible = false;
          CREATE_LOAD_SUMMARY_BUTTON.Visible = true;
          ADD_ALL_PANELS_BUTTON.Visible = true;
          if (GetPanelLoad() > 0) {
            ADD_ALL_PANELS_BUTTON.Enabled = false;
          }
          else {
            ADD_ALL_PANELS_BUTTON.Enabled = true;
          }
          SAFETY_FACTOR_CHECKBOX.Enabled = true;
        }
        else {
          if (i % 2 == 0) {
            PANEL_GRID.Rows[i].Cells["phase_a_left"].Style.BackColor = phaseColor;
            PANEL_GRID.Rows[i].Cells["phase_a_right"].Style.BackColor = phaseColor;
            PANEL_GRID.Rows[i].Cells["phase_b_left"].Style.BackColor = backColor1;
            PANEL_GRID.Rows[i].Cells["phase_b_right"].Style.BackColor = backColor1;
          }
          else {
            PANEL_GRID.Rows[i].Cells["phase_b_left"].Style.BackColor = phaseColor;
            PANEL_GRID.Rows[i].Cells["phase_b_right"].Style.BackColor = phaseColor;
            PANEL_GRID.Rows[i].Cells["phase_a_left"].Style.BackColor = backColor1;
            PANEL_GRID.Rows[i].Cells["phase_a_right"].Style.BackColor = backColor1;
          }
          PANEL_GRID.Rows[i].Cells["description_left"].Style.BackColor = backColor1;
          PANEL_GRID.Rows[i].Cells["description_right"].Style.BackColor = backColor1;
          PANEL_GRID.Rows[i].Cells["description_left"].Style.ForeColor = foreColor1;
          PANEL_GRID.Rows[i].Cells["description_right"].Style.ForeColor = foreColor1;
          PANEL_GRID.Rows[i].Cells["circuit_left"].Style.BackColor = backColor1;
          PANEL_GRID.Rows[i].Cells["circuit_right"].Style.BackColor = backColor1;
          PANEL_GRID.Rows[i].Cells["circuit_left"].Style.ForeColor = foreColor1;
          PANEL_GRID.Rows[i].Cells["circuit_right"].Style.ForeColor = foreColor1;
          PANEL_GRID.Rows[i].Cells["circuit_left"].ReadOnly = false;
          PANEL_GRID.Rows[i].Cells["circuit_right"].ReadOnly = false;
          PANEL_GRID.Rows[i].Cells["description_left"].ReadOnly = false;
          PANEL_GRID.Rows[i].Cells["description_right"].ReadOnly = false;
          PANEL_GRID.Rows[i].Cells["breaker_left"].Style.BackColor = backColor1;
          PANEL_GRID.Rows[i].Cells["breaker_right"].Style.BackColor = backColor1;
          PANEL_NAME_LABEL.Text = "PANEL";
          PANEL_NAME_LABEL.Location = new Point(150, 74);
          this.mainForm.PANEL_NAME_INPUT_TextChanged(sender, e, PANEL_NAME_INPUT.Text, false);
          CREATE_PANEL_BUTTON.Visible = true;
          CREATE_LOAD_SUMMARY_BUTTON.Visible = false;
          ADD_ALL_PANELS_BUTTON.Visible = false;
          SAFETY_FACTOR_CHECKBOX.Enabled = false;
        }
      }
    }

    public void ConfigureDistributionPanel(object sender, EventArgs e, bool updateCalcs = true) {
      if (this.is3Ph) {
        Color3pPanel(sender, e);
      }
      else {
        Color2pPanel(sender, e);
      }
      if (updateCalcs) SAFETY_FACTOR_CheckChanged(sender, e);
    }

    public void ClearModalAndRemoveRows(Dictionary<string, object> selectedPanelData) {
      ClearCurrentModalData();
      RemoveRows();

      int numberOfRows =
        ((Newtonsoft.Json.Linq.JArray)selectedPanelData["description_left"])
          .ToObject<List<string>>()
          .Count / 2;
      PANEL_GRID.Rows.Add(numberOfRows);
    }

    internal DataGridView RetrievePanelGrid() {
      return PANEL_GRID;
    }

    private void RemoveRows() {
      // remove rows
      while (PANEL_GRID.Rows.Count >= 1) {
        PANEL_GRID.Rows.RemoveAt(0);
      }
    }

    public void PopulateModalWithPanelData(Dictionary<string, object> selectedPanelData) {
      string GetSafeString(string key) {
        return selectedPanelData.TryGetValue(key, out object value) ? value?.ToString() ?? "" : "";
      }

      bool GetSafeBoolean(string key) {
        if (selectedPanelData.TryGetValue(key, out object value)) {
          if (value is bool boolValue) {
            return boolValue;
          }
          if (value is string stringValue) {
            return bool.TryParse(stringValue, out bool result) && result;
          }
        }
        return false;
      }

      // Set TextBoxes
      MAIN_INPUT.Text = GetSafeString("main").Replace("AMP", "").Replace("A", "").Replace(" ", "");
      PANEL_NAME_INPUT.Text = GetSafeString("panel").Replace("'", "");
      PANEL_LOCATION_INPUT.Text = GetSafeString("location");
      BUS_RATING_INPUT.Text = GetSafeString("bus_rating").Replace("AMP", "").Replace("A", "").Replace(" ", "");
      LCL.Text = GetSafeString("lcl");
      LML.Text = GetSafeString("lml");
      FED_FROM_TEXTBOX.Text = GetSafeString("fed_from");
      if (!String.IsNullOrEmpty(FED_FROM_TEXTBOX.Text)) {
        PHASE_COMBOBOX.Enabled = false;
        WIRE_COMBOBOX.Enabled = false;
      } else {
        PHASE_COMBOBOX.Enabled = true;
        WIRE_COMBOBOX.Enabled = true;
      }

      // Set Checkboxes
      LCL_OVERRIDE.Checked = GetSafeBoolean("lcl_override");
      LML_OVERRIDE.Checked = GetSafeBoolean("lml_override");
      DISTRIBUTION_SECTION_CHECKBOX.Checked = GetSafeBoolean("distribution_section");
      SAFETY_FACTOR_CHECKBOX.Checked = GetSafeBoolean("using_safety_factor");
      SAFETY_FACTOR_TEXTBOX.Text = GetSafeString("safety_factor");

      // Set ComboBoxes
      STATUS_COMBOBOX.SelectedItem = GetSafeString("existing");
      MOUNTING_COMBOBOX.SelectedItem = GetSafeString("mounting");
      WIRE_COMBOBOX.SelectedItem = GetSafeString("wire");
      PHASE_COMBOBOX.SelectedItem = GetSafeString("phase");
      LINE_VOLTAGE_COMBOBOX.SelectedItem = GetSafeString("voltage2");
      PHASE_VOLTAGE_COMBOBOX.SelectedItem = GetSafeString("voltage1");

      // Set DataGridViews
      PHASE_SUM_GRID.Rows[0].Cells[0].Value = GetSafeString("subtotal_a");
      PHASE_SUM_GRID.Rows[0].Cells[1].Value = GetSafeString("subtotal_b");
      if (PHASE_SUM_GRID.ColumnCount > 2) {
        PHASE_SUM_GRID.Rows[0].Cells[2].Value = GetSafeString("subtotal_c");
      }
      TOTAL_VA_GRID.Rows[0].Cells[0].Value = GetSafeString("total_va");
      PANEL_LOAD_GRID.Rows[0].Cells[0].Value = GetSafeString("kva");
      FEEDER_AMP_GRID.Rows[0].Cells[0].Value = GetSafeString("feeder_amps");
      if (SAFETY_FACTOR_CHECKBOX.Checked) {
        PANEL_LOAD_GRID.Rows[0].Cells[0].Value = (SafeConvertToDouble(GetSafeString("kva")) * SafeConvertToDouble(GetSafeString("safety_factor"))).ToString();
        FEEDER_AMP_GRID.Rows[0].Cells[0].Value = (SafeConvertToDouble(GetSafeString("feeder_amps")) * SafeConvertToDouble(GetSafeString("safety_factor"))).ToString();
      }

      // Set Custom Title if it exists
      if (selectedPanelData.TryGetValue("custom_title", out object customTitle)) {
        CUSTOM_TITLE_TEXT.Text = customTitle?.ToString() ?? "";
      }

      // Set id
      if (selectedPanelData.TryGetValue("id", out object outId)) {
        id = outId?.ToString() ?? System.Guid.NewGuid().ToString();
      }

      List<string> multi_row_datagrid_keys = new List<string>
      {
          "description_left",
          "description_right",
          "phase_a_left",
          "phase_b_left",
          "phase_a_right",
          "phase_b_right",
          "breaker_left",
          "breaker_right",
          "circuit_left",
          "circuit_right"
      };

      // Check if the panel is three phase and if so add the third phase to the list of keys
      if (selectedPanelData["phase"].ToString() == "3") {
        multi_row_datagrid_keys.AddRange(new List<string> { "phase_c_left", "phase_c_right" });
      }

      int length = ((Newtonsoft.Json.Linq.JArray)selectedPanelData["description_left"]).ToObject<List<string>>().Count;

      for (int i = 0; i < length * 2; i += 2) {
        foreach (string key in multi_row_datagrid_keys) {
          if (selectedPanelData[key] is Newtonsoft.Json.Linq.JArray) {
            List<string> values = ((Newtonsoft.Json.Linq.JArray)selectedPanelData[key]).ToObject<List<string>>();

            if (i < values.Count) {
              string currentValue = values[i];
              string nextValue = i + 1 < values.Count ? values[i + 1] : null;

              if (key.Contains("description") && currentValue == "SPACE") {
                currentValue = string.Empty;
              }

              if (key.Contains("phase") && currentValue == "0") {
                continue; // Skip this iteration if the value is "0" for phases
              }

              if (nextValue != null) {
                if (key.Contains("phase") && nextValue != "0") {
                  currentValue = $"{currentValue};{nextValue}";
                }
                else if (key.Contains("description") && nextValue != "SPACE") {
                  currentValue = $"{currentValue};{nextValue}";
                }
                else if (key.Contains("circuit")) {
                  currentValue = currentValue.Replace("A", "");
                }
                else if (key.Contains("breaker") && nextValue != "") {
                  currentValue = $"{currentValue};{nextValue}";
                }
              }

              // Check if PANEL_GRID has enough rows
              if (PANEL_GRID.Rows.Count <= i / 2) {
                Console.WriteLine($"Warning: PANEL_GRID does not have enough rows. Expected row: {i / 2}");
                continue;
              }

              // Check if PANEL_GRID has the specified column
              if (!PANEL_GRID.Columns.Contains(key)) {
                Console.WriteLine($"Warning: PANEL_GRID does not contain column: {key}");
                continue;
              }

              // Log values before assignment
              Console.WriteLine($"Setting PANEL_GRID.Rows[{i / 2}].Cells[{key}].Value = {currentValue}");

              // Check if the column index for the key is valid
              int columnIndex = PANEL_GRID.Columns[key].Index;
              if (columnIndex < 0 || columnIndex >= PANEL_GRID.ColumnCount) {
                Console.WriteLine($"Warning: Column index for {key} is out of range.");
                continue;
              }

              // Set the cell value
              PANEL_GRID.Rows[i / 2].Cells[columnIndex].Value = currentValue;
            }
            else {
              Console.WriteLine($"Warning: Index {i} is out of range for values in key {key}");
            }
          }
          else {
            // Log or handle the unexpected type
            Console.WriteLine($"Warning: Value for key {key} is not a JArray");
          }
        }
      }
      List<Dictionary<string, object>> panelData = this.mainForm.RetrieveSavedPanelData();
      for (int i = 0; i < PANEL_GRID.Rows.Count; i++) {
        if (IsUuid(PANEL_GRID.Rows[i].Cells["description_left"].Value as string)) {
          foreach (Dictionary<string, object> panel in panelData) {
            var panelId = panel["id"].ToString().Replace("\'", "").Replace("`", "");
            if (panelId.ToLower() == (PANEL_GRID.Rows[i].Cells["description_left"].Value as string).ToLower()) {
              string panelName = panel["panel"].ToString().Replace("\'", "").Replace("`", "");
              PANEL_GRID.Rows[i].Cells["description_left"].Value = "PANEL " + panelName;
            }
          }
        }
        if (IsUuid(PANEL_GRID.Rows[i].Cells["description_right"].Value as string)) {
          foreach (Dictionary<string, object> panel in panelData) {
            var panelId = panel["id"].ToString().Replace("\'", "").Replace("`", "");
            if (panelId.ToLower() == (PANEL_GRID.Rows[i].Cells["description_right"].Value as string).ToLower()) {
              string panelName = panel["panel"].ToString().Replace("\'", "").Replace("`", "");
              PANEL_GRID.Rows[i].Cells["description_right"].Value = "PANEL " + panelName;
            }
          }
        }
      }
      isLoading = false;
    }

    private void ClearCurrentModalData() {
      // Clear TextBoxes
      BUS_RATING_INPUT.Text = string.Empty;
      MAIN_INPUT.Text = string.Empty;
      PANEL_LOCATION_INPUT.Text = string.Empty;
      PANEL_NAME_INPUT.Text = string.Empty;
      LCL.Text = "0";
      LML.Text = "0";

      // Clear ComboBoxes
      STATUS_COMBOBOX.SelectedIndex = -1; // This will unselect all items
      MOUNTING_COMBOBOX.SelectedIndex = -1;
      WIRE_COMBOBOX.SelectedIndex = -1;
      PHASE_COMBOBOX.SelectedIndex = -1;
      LINE_VOLTAGE_COMBOBOX.SelectedIndex = -1;
      PHASE_VOLTAGE_COMBOBOX.SelectedIndex = -1;

      // Clear DataGridViews
      PHASE_SUM_GRID.Rows[0].Cells[0].Value = "0";
      PHASE_SUM_GRID.Rows[0].Cells[1].Value = "0";
      TOTAL_VA_GRID.Rows[0].Cells[0].Value = "0";
      PANEL_LOAD_GRID.Rows[0].Cells[0].Value = "0";
      FEEDER_AMP_GRID.Rows[0].Cells[0].Value = "0";

      // Clear DataGridViews
      for (int i = 0; i < PANEL_GRID.Rows.Count; i++) {
        PANEL_GRID.Rows[i].Cells["description_left"].Value = string.Empty;
        PANEL_GRID.Rows[i].Cells["description_right"].Value = string.Empty;
        PANEL_GRID.Rows[i].Cells["phase_a_left"].Value = string.Empty;
        PANEL_GRID.Rows[i].Cells["phase_b_left"].Value = string.Empty;
        PANEL_GRID.Rows[i].Cells["phase_a_right"].Value = string.Empty;
        PANEL_GRID.Rows[i].Cells["phase_b_right"].Value = string.Empty;
        PANEL_GRID.Rows[i].Cells["breaker_left"].Value = string.Empty;
        PANEL_GRID.Rows[i].Cells["breaker_right"].Value = string.Empty;
        PANEL_GRID.Rows[i].Cells["circuit_left"].Value = string.Empty;
        PANEL_GRID.Rows[i].Cells["circuit_right"].Value = string.Empty;
      }
    }

    private void AddPhaseSumColumn(bool is3Ph) {
      if (is3Ph) {
        PHASE_SUM_GRID.Columns.Add(PHASE_SUM_GRID.Columns[0].Clone() as DataGridViewColumn);
        PHASE_SUM_GRID.Columns[2].HeaderText = "PH C (VA)";
        PHASE_SUM_GRID.Columns[2].Name = "TOTAL_PH_C";

        // Set the width of the new column
        PHASE_SUM_GRID.Columns[2].Width = 80;

        // Set the width of the other columns
        PHASE_SUM_GRID.Columns[0].Width = 80;
        PHASE_SUM_GRID.Columns[1].Width = 80;

        // set the width of the grid
        PHASE_SUM_GRID.Width = 285;
        PHASE_SUM_GRID.Location = new System.Drawing.Point(12, 335);
      }
      else {
        if (PHASE_SUM_GRID.Columns.Count > 2) {
          PHASE_SUM_GRID.Columns.Remove("TOTAL_PH_C");
        }

        // Set the width of the other columns
        PHASE_SUM_GRID.Columns[0].Width = 100;
        PHASE_SUM_GRID.Columns[1].Width = 100;

        // set the width of the grid
        PHASE_SUM_GRID.Width = 245;
        PHASE_SUM_GRID.Location = new System.Drawing.Point(52, 335);
      }
    }

    private void LinkCellToPhase(string cellValue, DataGridViewRow row, DataGridViewColumn col) {
      var (panel_name, phase) = ConvertCellValueToPanelNameAndPhase(cellValue);
      if (panel_name.ToLower() == PANEL_NAME_INPUT.Text.ToLower()) {
        return;
      }

      var isPanelReal = this.mainForm.PanelNameExists(panel_name);

      if (isPanelReal) {
        UserControl panelControl = mainForm.FindUserControl(panel_name);

        if (panelControl != null) {
          DataGridView panelControl_phaseSumGrid =
            panelControl.Controls.Find("PHASE_SUM_GRID", true).FirstOrDefault() as DataGridView;
          DataGridView this_panelGrid =
            this.Controls.Find("PANEL_GRID", true).FirstOrDefault() as DataGridView;
          this_panelGrid.Rows[row.Index].Cells[col.Index].Tag = cellValue;
          ListenForPhaseChanges(panelControl_phaseSumGrid, phase, row, col, this_panelGrid);
        }
      }
    }

    private void ListenForPhaseChanges(
      DataGridView panelControl_phaseSumGrid,
      string phase,
      DataGridViewRow row,
      DataGridViewColumn col,
      DataGridView panelGrid
    ) {
      var phaseSumGrid_row = 0;
      var phaseSumGrid_col = 0;

      DataGridViewCellEventHandler eventHandler = null;
      DataGridViewCellEventHandler panelGrid_eventHandler = null;

      if (phase == "A") {
        phaseSumGrid_col = 0;
      }
      else if (phase == "B") {
        phaseSumGrid_col = 1;
      }
      else if (phase == "C") {
        phaseSumGrid_col = 2;
      }

      var newCellValue = panelControl_phaseSumGrid
        .Rows[phaseSumGrid_row]
        .Cells[phaseSumGrid_col]
        .Value.ToString();
      panelGrid.Rows[row.Index].Cells[col.Index].Value = newCellValue;
      panelGrid.Rows[row.Index].Cells[col.Index].Style.BackColor = Color.LightGreen;

      eventHandler = (sender, e) => {
        if (e.RowIndex == phaseSumGrid_row && e.ColumnIndex == phaseSumGrid_col) {
          var newCellValue = panelControl_phaseSumGrid
            .Rows[e.RowIndex]
            .Cells[e.ColumnIndex]
            .Value.ToString();
          panelGrid.Rows[row.Index].Cells[col.Index].Value = newCellValue;
          panelGrid.Rows[row.Index].Cells[col.Index].Style.BackColor = Color.LightGreen;
        }
      };

      panelControl_phaseSumGrid.CellValueChanged += eventHandler;

      panelGrid_eventHandler = (sender, e) => {
        if (e.RowIndex == row.Index && e.ColumnIndex == col.Index) {
          var newCellValue = panelGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString();
          try {
            var phaseSumGridValue = panelControl_phaseSumGrid
            .Rows[phaseSumGrid_row]
            .Cells[phaseSumGrid_col]
            .Value?.ToString();
            if (newCellValue != phaseSumGridValue) {
              panelControl_phaseSumGrid.CellValueChanged -= eventHandler;
              panelGrid.CellValueChanged -= panelGrid_eventHandler;
              panelGrid.Rows[row.Index].Cells[col.Index].Style.BackColor = Color.LightGray;
              if (panelGrid.Rows[row.Index].Cells[col.Index].Tag != null) {
                panelGrid.Rows[row.Index].Cells[col.Index].Tag = null;
              }
            }
          }
          catch (Exception ex) {
            panelControl_phaseSumGrid.CellValueChanged -= eventHandler;
            panelGrid.CellValueChanged -= panelGrid_eventHandler;
            panelGrid.Rows[row.Index].Cells[col.Index].Style.BackColor = Color.LightGray;
            if (panelGrid.Rows[row.Index].Cells[col.Index].Tag != null) {
              panelGrid.Rows[row.Index].Cells[col.Index].Tag = null;
            }
          }
        }
      };

      panelGrid.CellValueChanged += panelGrid_eventHandler;
    }

    private (string, string) ConvertCellValueToPanelNameAndPhase(string cellValue) {
      cellValue = cellValue.ToUpper();
      Regex regex = new Regex(@"^=[a-zA-Z0-9]*-[A-C]$");
      if (!regex.IsMatch(cellValue)) {
        return ("", "");
      }

      string[] splitCellValue = cellValue.Split('-');
      string panelName = splitCellValue[0].Replace("=", "");
      string phase = splitCellValue[1];

      return (panelName, phase);
    }

    private void ChangeSizeOfPhaseColumns(bool is3Ph) {
      if (is3Ph) {
        // Left Side
        PANEL_GRID.Columns["phase_a_left"].Width = 67;
        PANEL_GRID.Columns["phase_b_left"].Width = 67;
        PANEL_GRID.Columns["phase_c_left"].Width = 67;

        // Right Side
        PANEL_GRID.Columns["phase_a_right"].Width = 67;
        PANEL_GRID.Columns["phase_b_right"].Width = 67;
        PANEL_GRID.Columns["phase_c_right"].Width = 67;
      }
      else {
        // Left Side
        PANEL_GRID.Columns["phase_a_left"].Width = 100;
        PANEL_GRID.Columns["phase_b_left"].Width = 100;

        // Right Side
        PANEL_GRID.Columns["phase_a_right"].Width = 100;
        PANEL_GRID.Columns["phase_b_right"].Width = 100;
      }
    }

    private void AddOrRemovePanelGridColumns(bool is3Ph) {
      if (is3Ph) {
        // Left Side
        DataGridViewTextBoxColumn phase_c_left = new DataGridViewTextBoxColumn();
        phase_c_left.HeaderText = "PH C";
        phase_c_left.Name = "phase_c_left";
        phase_c_left.Width = 50;
        PANEL_GRID.Columns.Insert(3, phase_c_left);

        // Right Side
        DataGridViewTextBoxColumn phase_c_right = new DataGridViewTextBoxColumn();
        phase_c_right.HeaderText = "PH C";
        phase_c_right.Name = "phase_c_right";
        phase_c_right.Width = 50;
        PANEL_GRID.Columns.Insert(10, phase_c_right);
      }
      else {
        if (PANEL_GRID.Columns.Count > 10) {
          // Left Side
          PANEL_GRID.Columns.Remove("phase_c_left");

          // Right Side
          PANEL_GRID.Columns.Remove("phase_c_right");
        }
      }
    }

    private void UpdateApplyComboboxToMatchStorage() {
      var apply_combobox_items = new List<string>();
      foreach (var note in this.notesStorage) {
        if (!apply_combobox_items.Contains(note)) {
          apply_combobox_items.Add(note);
        }
      }

      APPLY_COMBOBOX.DataSource = apply_combobox_items;
    }

    private void RemoveTagsFromCells(string tag) {
      foreach (DataGridViewRow row in PANEL_GRID.Rows) {
        foreach (DataGridViewCell cell in row.Cells) {
          if (cell.Tag != null) {
            string cellTag = cell.Tag.ToString();
            if (cellTag.Contains(tag)) {
              cellTag = cellTag.Replace(tag, "");
              if (cellTag.EndsWith("|")) {
                cellTag = cellTag.Substring(0, cellTag.Length - 1);
              }
              cell.Tag = cellTag;
            }
          }
        }
      }
    }

    public void UpdateCellBackgroundColor() {
      if (APPLY_COMBOBOX.SelectedItem == null) {
        return;
      }

      foreach (DataGridViewRow row in PANEL_GRID.Rows) {
        foreach (DataGridViewCell cell in row.Cells) {
          if (cell.OwningColumn.Name.Contains("description")) {
            cell.Style.BackColor = Color.Empty;
          }
          else if (cell.OwningColumn.Name.Contains("phase") && cell.Tag != null && cell.Tag.ToString().Contains("LCL125")) {
            cell.Style.BackColor = Color.Salmon;
          }
          else if (cell.OwningColumn.Name.Contains("phase") && cell.Tag != null && cell.Tag.ToString().Contains("LCL80")) {
            cell.Style.BackColor = Color.Gold;
          }
        }
      }

      foreach (DataGridViewRow row in PANEL_GRID.Rows) {
        foreach (DataGridViewCell cell in row.Cells) {
          if (cell.OwningColumn.Name.Contains("description")) {
            if (cell.Tag == null) {
              continue;
            }
            if (cell.Tag.ToString().Split('|').Contains(APPLY_COMBOBOX.SelectedItem.ToString())) {
              // turn the background of the cell to a yellow color
              cell.Style.BackColor = Color.Yellow;
            }
          }
        }
      }
    }

    public void UpdateNotesStorage(List<string> notesStorage) {
      this.notesStorage = notesStorage;
      UpdateApplyComboboxToMatchStorage();
    }

    private void UpdatePanelGrid(
      int phaseSumGridColumnCount,
      int panelPhaseSumGridColumnCount,
      int rowIndex,
      string side,
      string panelName
    ) {
      if (phaseSumGridColumnCount == panelPhaseSumGridColumnCount) {
        int rowCount = phaseSumGridColumnCount == 3 ? 3 : 2;
        if (PANEL_GRID.Rows.Count > rowIndex + rowCount - 1) {
          var cellValueA = "=" + panelName.ToUpper() + "-A";
          var cellValueB = "=" + panelName.ToUpper() + "-B";
          var cellValueC = phaseSumGridColumnCount == 3 ? "=" + panelName.ToUpper() + "-C" : null;

          List<DataGridViewRow> rows = new List<DataGridViewRow>();
          for (int i = 0; i < rowCount; i++) {
            rows.Add(PANEL_GRID.Rows[rowIndex + i]);
          }

          List<string> phases = new List<string> { "phase_a_", "phase_b_" };
          if (phaseSumGridColumnCount == 3) {
            phases.Add("phase_c_");
          }

          List<string> cellValues = new List<string> { cellValueA, cellValueB };
          if (cellValueC != null) {
            cellValues.Add(cellValueC);
          }

          for (int i = 0; i < phases.Count; i++) {
            foreach (DataGridViewRow gridRow in rows) {
              string cellName = phases[i] + side;
              if (
                gridRow.Cells[cellName].Style.BackColor == Color.LightGray
                || gridRow.Cells[cellName].Style.BackColor == Color.LightGreen
              ) {
                gridRow.Cells[cellName].Value = cellValues[i];
                if (i == phases.Count - 1) {
                  gridRow.Cells[$"breaker_{side}"].Value = phases.Count.ToString();
                }
                else if (phases.Count == 3 && i == 1) {
                  gridRow.Cells[$"breaker_{side}"].Value = "";
                }
              }
            }       
          }
        }
      }
      else if (phaseSumGridColumnCount == 2 && panelPhaseSumGridColumnCount == 3) {
        if (PANEL_GRID.Rows.Count > rowIndex + 1) {
          var phases = new List<string> { "A", "B" };
          if (side == "left") {
            for (int i = rowIndex; i < rowIndex + 2; i++) // Loop for the first two rows
            {
              foreach (string colName in new[] { "phase_a_left", "phase_b_left", "phase_c_left" }) // Loop for the specified columns
              {
                var cell = PANEL_GRID.Rows[i].Cells[colName];
                if (
                  cell.Style.BackColor == Color.LightGray
                  || cell.Style.BackColor == Color.LightGreen
                ) // Check the background color
                {
                  cell.Value = "=" + panelName + "-" + phases[i - rowIndex];
                }
              }
            }
          }
          else {
            for (int i = rowIndex; i < rowIndex + 2; i++) // Loop for the first two rows
            {
              foreach (
                string colName in new[] { "phase_a_right", "phase_b_right", "phase_c_right" }
              ) // Loop for the specified columns
              {
                var cell = PANEL_GRID.Rows[i].Cells[colName];
                if (
                  cell.Style.BackColor == Color.LightGray
                  || cell.Style.BackColor == Color.LightGreen
                ) // Check the background color
                {
                  cell.Value = "=" + panelName + "-" + phases[i - rowIndex];
                }
              }
            }
          }
        }
      }
    }

    private bool BreakerContainsNote(int rowIndex, string columnName, string note) {
      string columnPrefix = columnName.Contains("left") ? "left" : "right";
      var descriptionCellTag = PANEL_GRID.Rows[rowIndex].Cells[$"description_{columnPrefix}"].Tag;
      var descriptionCellValue = PANEL_GRID
        .Rows[rowIndex]
        .Cells[$"description_{columnPrefix}"]
        .Value;
      var breakerCellValue = PANEL_GRID.Rows[rowIndex].Cells[$"breaker_{columnPrefix}"].Value;

      if (descriptionCellTag != null && descriptionCellTag.ToString().Contains(note)) {
        return true;
      }

      if (
        descriptionCellValue == null
        || descriptionCellValue.ToString() == ""
        || descriptionCellValue.ToString().All(c => c == '-')
      ) {
        if (
          breakerCellValue != null
          && (breakerCellValue.ToString() == "2" || breakerCellValue.ToString() == "3")
        ) {
          int rowsAbove = breakerCellValue.ToString() == "2" ? 1 : 2;
          if (rowIndex < rowsAbove) {
            return false;
          }
          var descriptionCellTagAbove = PANEL_GRID
            .Rows[rowIndex - rowsAbove]
            .Cells[$"description_{columnPrefix}"]
            .Tag;
          if (descriptionCellTagAbove != null && descriptionCellTagAbove.ToString().Contains(note)) {
            return true;
          }
        }

        if (breakerCellValue == null || breakerCellValue.ToString() == "") {
          if (rowIndex == PANEL_GRID.Rows.Count - 1) {
            return false;
          }
          var nextBreakerCellValue = PANEL_GRID
            .Rows[rowIndex + 1]
            .Cells[$"breaker_{columnPrefix}"]
            .Value;
          if (nextBreakerCellValue != null && nextBreakerCellValue.ToString() == "3") {
            var descriptionCellTagAbove = PANEL_GRID
              .Rows[rowIndex - 1]
              .Cells[$"description_{columnPrefix}"]
              .Tag;
            if (
              descriptionCellTagAbove != null
              && descriptionCellTagAbove.ToString().Contains(note)
            ) {
              return true;
            }
          }
        }
      }
      return false;
    }

    private string CalculateCellOrLinkPanel(
      DataGridViewCellEventArgs e,
      string cellValue,
      DataGridViewRow row,
      DataGridViewColumn col
    ) {
      if (cellValue.StartsWith("=")) {
        if (col.Name.Contains("phase")) {
          cellValue = cellValue.Replace(" ", "");
        }
        if (
          cellValue.All(c =>
            char.IsDigit(c)
            || c == '.'
            || c == '='
            || c == '-'
            || c == '*'
            || c == '+'
            || c == '/'
            || c == '('
            || c == ')'
          )
        ) {
          var result = new System.Data.DataTable().Compute(cellValue.Replace("=", ""), null);
          PANEL_GRID.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = Math.Ceiling(Convert.ToDouble(result));
        }
        else {
          LinkCellToPhase(cellValue, row, col);
        }
      }

      return cellValue;
    }

    private void AutoSetBreakerSize(string cellValue, DataGridViewRow row, DataGridViewColumn col) {
      string side = "left";
      for (int i = 0; i < 2; i++) {
        if (!String.IsNullOrEmpty(row.Cells[$"description_{side}"].Value as string) &&
             (
               !String.IsNullOrEmpty(row.Cells[$"breaker_{side}"].Value as string) &&
               row.Cells[$"breaker_{side}"].Value as string == "20"
             )
           ) {
          if (col.Name.Contains("phase_") && col.Name.Contains(side)) {
            if (double.TryParse(cellValue, out double val)) {
              if (double.TryParse(PHASE_VOLTAGE_COMBOBOX.Text, out double pv)) {
                if (val / pv * 1.25 > 20) {
                  double mocp = Math.Round(val / pv * 1.25, 0);
                  switch (mocp) {
                    case var _ when mocp <= 25:
                      row.Cells[$"breaker_{side}"].Value = "25";
                      break;
                    case var _ when mocp <= 30:
                      row.Cells[$"breaker_{side}"].Value = "30";
                      break;
                    case var _ when mocp <= 35:
                      row.Cells[$"breaker_{side}"].Value = "35";
                      break;
                    case var _ when mocp <= 40:
                      row.Cells[$"breaker_{side}"].Value = "40";
                      break;
                    case var _ when mocp <= 45:
                      row.Cells[$"breaker_{side}"].Value = "45";
                      break;
                    case var _ when mocp <= 50:
                      row.Cells[$"breaker_{side}"].Value = "50";
                      break;
                    case var _ when mocp <= 60:
                      row.Cells[$"breaker_{side}"].Value = "60";
                      break;
                    case var _ when mocp <= 70:
                      row.Cells[$"breaker_{side}"].Value = "70";
                      break;
                    case var _ when mocp <= 80:
                      row.Cells[$"breaker_{side}"].Value = "80";
                      break;
                    case var _ when mocp <= 90:
                      row.Cells[$"breaker_{side}"].Value = "90";
                      break;
                    case var _ when mocp <= 100:
                      row.Cells[$"breaker_{side}"].Value = "100";
                      break;
                    case var _ when mocp <= 110:
                      row.Cells[$"breaker_{side}"].Value = "110";
                      break;
                    case var _ when mocp <= 125:
                      row.Cells[$"breaker_{side}"].Value = "125";
                      break;
                    case var _ when mocp <= 150:
                      row.Cells[$"breaker_{side}"].Value = "150";
                      break;
                    case var _ when mocp <= 175:
                      row.Cells[$"breaker_{side}"].Value = "175";
                      break;
                    case var _ when mocp <= 200:
                      row.Cells[$"breaker_{side}"].Value = "200";
                      break;
                    case var _ when mocp <= 225:
                      row.Cells[$"breaker_{side}"].Value = "225";
                      break;
                    case var _ when mocp <= 250:
                      row.Cells[$"breaker_{side}"].Value = "250";
                      break;
                    case var _ when mocp <= 300:
                      row.Cells[$"breaker_{side}"].Value = "300";
                      break;
                    case var _ when mocp <= 350:
                      row.Cells[$"breaker_{side}"].Value = "350";
                      break;
                    case var _ when mocp <= 400:
                      row.Cells[$"breaker_{side}"].Value = "400";
                      break;
                    case var _ when mocp <= 450:
                      row.Cells[$"breaker_{side}"].Value = "450";
                      break;
                    case var _ when mocp <= 500:
                      row.Cells[$"breaker_{side}"].Value = "500";
                      break;
                    case var _ when mocp <= 600:
                      row.Cells[$"breaker_{side}"].Value = "600";
                      break;
                  }
                }
              }
            }
          }
          side = "right";
        }
      }
    }

    public void LinkSubpanels() {
      string side = "left";
      for (int j = 0; j < 2; j++) {
        for (int i = 0; i < PANEL_GRID.Rows.Count; i++) {
          if (PANEL_GRID.Rows[i].Cells[$"description_{side}"].Value.ToString().ToUpper().Contains("PANEL")) {
            LinkSubpanel(PANEL_GRID.Rows[i].Cells[$"description_{side}"].Value.ToString(), i, side);
          }
        }
        side = "right";
      }
    }

    private void LinkSubpanel(string cellValue, int rowIndex, string side) {
      var panelName = cellValue.ToLower().Split(' ').Last();

      if (panelName.Contains("'") || panelName.Contains("`")) {
        panelName = panelName.Replace("'", "").Replace("`", "");
      }

      if (panelName.ToUpper() == PANEL_NAME_INPUT.Text.ToUpper()) return;

      var isPanelReal = this.mainForm.PanelNameExists(panelName);

      if (isPanelReal) {
        PanelUserControl panelControl = (PanelUserControl)mainForm.FindUserControl(panelName);
        DataGridView panelControl_phaseSumGrid =
          panelControl.Controls.Find("PHASE_SUM_GRID", true).FirstOrDefault() as DataGridView;
        UpdateSubpanelFedFrom();
        var phaseSumGridColumnCount = panelControl_phaseSumGrid.ColumnCount;
        var panelPhaseSumGridColumnCount = PHASE_SUM_GRID.ColumnCount;
        
        UpdatePanelGrid(
          phaseSumGridColumnCount,
          panelPhaseSumGridColumnCount,
          rowIndex,
          side,
          panelName
        );
        UpdatePerCellValueChange();
      }
    }

    private void AutoLinkSubpanels(string cellValue, DataGridViewRow row, DataGridViewColumn col) {
      if (col.Name.Contains("description")) {
        if (cellValue.ToUpper().Contains("PANEL")) {
          string side = col.Name.Contains("left") ? "left" : "right";
          int rowIndex = row.Index;
          LinkSubpanel(cellValue, rowIndex, side);
        }
      }
    }

    private void RemoveExistingBreakerNote(DataGridViewCell dataGridViewCell) {
      if (!dataGridViewCell.OwningColumn.Name.Contains("breaker")) {
        return;
      }

      var side = dataGridViewCell.OwningColumn.Name.Contains("left") ? "left" : "right";

      if (PANEL_GRID.Rows[dataGridViewCell.RowIndex].Cells["description_" + side].Tag == null) {
        return;
      }

      var descriptionCellTag = PANEL_GRID
        .Rows[dataGridViewCell.RowIndex]
        .Cells["description_" + side]
        .Tag;
      descriptionCellTag = descriptionCellTag
        .ToString()
        .Replace("DENOTES EXISTING CIRCUIT BREAKER TO REMAIN; ALL OTHERS ARE NEW.", "");
      descriptionCellTag = descriptionCellTag.ToString().TrimEnd('|');

      PANEL_GRID.Rows[dataGridViewCell.RowIndex].Cells["description_" + side].Tag =
        descriptionCellTag;

      if (
        APPLY_COMBOBOX.SelectedItem != null
        && APPLY_COMBOBOX.SelectedItem.ToString()
          == "DENOTES EXISTING CIRCUIT BREAKER TO REMAIN; ALL OTHERS ARE NEW."
      ) {
        PANEL_GRID.Rows[dataGridViewCell.RowIndex].Cells["description_" + side].Style.BackColor =
          Color.White;
      }
    }

    private void RemoveExistingFromDescription(DataGridViewCell dataGridViewCell) {
      if (!dataGridViewCell.OwningColumn.Name.Contains("description")) {
        return;
      }

      if (dataGridViewCell.Tag == null) {
        return;
      }

      var cellTag = dataGridViewCell.Tag.ToString();
      cellTag = cellTag.Replace("ADD SUFFIX (E). *NOT ADDED AS NOTE*", "");
      cellTag = cellTag.ToString().TrimEnd('|');

      dataGridViewCell.Tag = cellTag;

      if (
        APPLY_COMBOBOX.SelectedItem != null
        && APPLY_COMBOBOX.SelectedItem.ToString() == "ADD SUFFIX (E). *NOT ADDED AS NOTE*"
      ) {
        dataGridViewCell.Style.BackColor = Color.White;
      }
    }

    private void PANEL_NAME_INPUT_TextChanged(object sender, EventArgs e) {
      this.mainForm.PANEL_NAME_INPUT_TextChanged(sender, e, PANEL_NAME_INPUT.Text.ToUpper(), DISTRIBUTION_SECTION_CHECKBOX.Checked);
      this.Name = PANEL_NAME_INPUT.Text.ToUpper();
    }

    private void PANEL_NAME_INPUT_Leave(object sender, EventArgs e) {
      this.mainForm.PANEL_NAME_INPUT_Leave(sender, e, PANEL_NAME_INPUT.Text.ToUpper(), id, FED_FROM_TEXTBOX.Text);
    }

    private void PANEL_GRID_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e) {
      if (PHASE_SUM_GRID.ColumnCount > 2) {
        ListenFor3pRowsAdded(e);
      }
      else {
        ListenFor2pRowsAdded(e);
      }
      ConfigureDistributionPanel(sender, e, false);
    }

    private void DISTRIBUTION_SECTION_CHECKBOX_Checked(object sender, EventArgs e) {
      ConfigureDistributionPanel(sender, e);
    }

    private void PANEL_GRID_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e) {
      e.CellStyle.SelectionBackColor = e.CellStyle.BackColor;
      e.CellStyle.SelectionForeColor = e.CellStyle.ForeColor;
    }

    private void PANEL_GRID_CellValueChanged(object sender, DataGridViewCellEventArgs e) {
      if (e.RowIndex < 0 || e.ColumnIndex < 0) {
        return;
      }
      RemoveExistingFromDescription(PANEL_GRID.Rows[e.RowIndex].Cells[e.ColumnIndex]);
      RemoveExistingBreakerNote(PANEL_GRID.Rows[e.RowIndex].Cells[e.ColumnIndex]);

      if (PANEL_GRID.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null) {
        return;
      }

      CalculateBreakerLoad();
      UpdatePerCellValueChange();
    }

    private void PANEL_GRID_CellValueChangedLink(object sender, DataGridViewCellEventArgs e) {
      var cell = PANEL_GRID.Rows[e.RowIndex].Cells[e.ColumnIndex];
      if (cell == null) {
        return;
      }

      string cellValue = cell.Value?.ToString() ?? string.Empty;
      if (!String.IsNullOrEmpty(cellValue)) {
        cell.Value = cell.Value.ToString().ToUpper();
      }
      var row = PANEL_GRID.Rows[e.RowIndex];
      var col = PANEL_GRID.Columns[e.ColumnIndex];

      if (row != null && col != null) {
        AutoLinkSubpanels(cellValue, row, col);
        AutoSetBreakerSize(cellValue, row, col);
        cellValue = CalculateCellOrLinkPanel(e, cellValue, row, col);
      }
    }

    private void PANEL_GRID_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e) {
      selectedCells = new List<DataGridViewCell>(PANEL_GRID.SelectedCells.Cast<DataGridViewCell>());
      oldValue = PANEL_GRID.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
    }

    private void PANEL_GRID_KeyDown(object sender, KeyEventArgs e) {
      if (e.Control && e.KeyCode == Keys.V) {
        // Get text from clipboard
        string text = Clipboard.GetText();

        if (!string.IsNullOrEmpty(text)) {
          // Split clipboard text into lines
          string[] lines = text.Split('\n');

          if (lines.Length > 0 && string.IsNullOrWhiteSpace(lines[lines.Length - 1])) {
            Array.Resize(ref lines, lines.Length - 1);
          }

          // Get start cell for pasting
          int rowIndex = PANEL_GRID.CurrentCell.RowIndex;
          int colIndex = PANEL_GRID.CurrentCell.ColumnIndex;

          // Paste each line into a row
          foreach (string line in lines) {
            string[] parts = line.Split('\t');

            for (int i = 0; i < parts.Length; i++) {
              if (rowIndex < PANEL_GRID.RowCount && colIndex + i < PANEL_GRID.ColumnCount) {
                try {
                  PANEL_GRID[colIndex + i, rowIndex].Value = parts[i].Trim();
                }
                catch (FormatException) {
                  // Set to default value
                  PANEL_GRID[colIndex + i, rowIndex].Value = 0;

                  // Or notify user
                  MessageBox.Show("Invalid format in cell!");
                }
              }
            }

            rowIndex++;
          }
        }

        e.Handled = true;
      }
      // Check if Ctrl+C was pressed
      else if (e.Control && e.KeyCode == Keys.C) {
        StringBuilder copiedText = new StringBuilder();

        // Get the minimum and maximum rowIndex and columnIndex of the selected cells
        int minRowIndex = PANEL_GRID
          .SelectedCells.Cast<DataGridViewCell>()
          .Min(cell => cell.RowIndex);
        int maxRowIndex = PANEL_GRID
          .SelectedCells.Cast<DataGridViewCell>()
          .Max(cell => cell.RowIndex);
        int minColumnIndex = PANEL_GRID
          .SelectedCells.Cast<DataGridViewCell>()
          .Min(cell => cell.ColumnIndex);
        int maxColumnIndex = PANEL_GRID
          .SelectedCells.Cast<DataGridViewCell>()
          .Max(cell => cell.ColumnIndex);

        // Loop through the rows
        for (int rowIndex = minRowIndex; rowIndex <= maxRowIndex; rowIndex++) {
          List<string> cellValues = new List<string>();

          // Loop through the columns
          for (int columnIndex = minColumnIndex; columnIndex <= maxColumnIndex; columnIndex++) {
            DataGridViewCell cell = PANEL_GRID[columnIndex, rowIndex];

            // Only add the cell value to the list if the cell is selected
            if (cell.Selected) {
              cellValues.Add(cell.Value?.ToString() ?? string.Empty);
            }
          }

          // Add the cell values of the row to the copied text
          if (cellValues.Count > 0) {
            copiedText.AppendLine(string.Join("\t", cellValues));
          }
        }

        if (copiedText.Length > 0) {
          Clipboard.SetText(copiedText.ToString());
        }

        e.Handled = true;
      }
      // Existing code for handling the Delete key
      else if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back) {
        foreach (DataGridViewCell cell in PANEL_GRID.SelectedCells) {
          if (!String.IsNullOrEmpty(cell.Value as string) && (cell.Value as string).ToUpper().Contains("PANEL")) {
            this.mainForm.RemoveFedFrom(cell.Value as string);
          }
          cell.Value = "";
          cell.Tag = null;
        }
        e.Handled = true;

        UpdatePerCellValueChange();
      }
      // Check if Ctrl+D was pressed
      else if (e.Control && e.KeyCode == Keys.D) {
        int rowIndex = PANEL_GRID.CurrentCell.RowIndex;
        int colIndex = PANEL_GRID.CurrentCell.ColumnIndex;

        // Check if there is a row above
        if (rowIndex > 0) {
          // Get the value from the cell above
          object value = PANEL_GRID[colIndex, rowIndex - 1].Value;

          // Paste the value into the current cell
          PANEL_GRID[colIndex, rowIndex].Value = value;
        }

        e.Handled = true;
      }
    }

    private void SAFETY_FACTOR_TEXTBOX_KeyUp(object sender, KeyEventArgs e) {
      if (!Regex.IsMatch(SAFETY_FACTOR_TEXTBOX.Text, @"^\d*\.?\d*$")) {
        SAFETY_FACTOR_TEXTBOX.BackColor = Color.Crimson;
      }
      else {
        SAFETY_FACTOR_TEXTBOX.BackColor = Color.White;
        UpdatePerCellValueChange();
      }
    }

    private void BUS_RATING_INPUT_TEXTBOX_KeyUp(object sender, KeyEventArgs e) {
      if (!Regex.IsMatch(BUS_RATING_INPUT.Text, @"^\d*\.?\d*$")) {
        BUS_RATING_INPUT.BackColor = Color.Crimson;
      }
      else if (!String.IsNullOrEmpty(BUS_RATING_INPUT.Text)) {
        BUS_RATING_INPUT.BackColor = Color.White;
        UpdatePerCellValueChange();
      }
    }

    private async void PANEL_GRID_CellClick(object sender, DataGridViewCellEventArgs e) {
      if (e.RowIndex == -1) {
        return;
      }

      if (e.ColumnIndex < 0) {
        return;
      }

      if (!PANEL_GRID.Columns[e.ColumnIndex].Name.Contains("phase")) {
        return;
      }

      // Get the selected cell
      DataGridViewCell cell = PANEL_GRID.Rows[e.RowIndex].Cells[e.ColumnIndex];

      // check if the cell has a tag
      if (cell.Tag != null) {
        if (!cell.Tag.ToString().Contains("=")) return;
        // remove the equals from the tag
        string cellValue = cell.Tag.ToString().Replace("=", "");

        // split the cell value by dash, to create two strings, one for the panel name and one for the phase
        string[] splitCellValue = cellValue.Split('-');

        // get the panel name and phase from the split cell value
        string panelName = splitCellValue[0];
        string phase = splitCellValue[1];

        // change the INFO_LABEL.TEXT value to inform the user that the cell is linked to another panel and its phase for 5 seconds, then erase it
        INFO_LABEL.Text = $"This cell is linked to phase {phase} of panel '{panelName}'.";

        // wait for 5 seconds
        await Task.Delay(5000);

        // if the INFO_LABEL.TEXT value is still the same as it was 5 seconds ago, then erase it
        if (INFO_LABEL.Text == $"This cell is linked to phase {phase} of panel '{panelName}'.") {
          INFO_LABEL.Text = string.Empty;
        }
      }
    }

    private void PANEL_GRID_CellPainting(object sender, DataGridViewCellPaintingEventArgs e) {
      if (e.RowIndex >= 0 && e.ColumnIndex >= 0) // Check if it's not the header cell
      {
        var cell = PANEL_GRID[e.ColumnIndex, e.RowIndex];
        if (cell.Selected) {
          e.Paint(e.CellBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.Border);

          using (Pen p = new Pen(Color.Black, 2)) // Change to desired border color and size
          {
            Rectangle rect = e.CellBounds;
            rect.Width -= 2;
            rect.Height -= 2;
            e.Graphics.DrawRectangle(p, rect);
          }

          e.Handled = true;
        }
      }
    }

    private double AggregateSinglePhaseLoads() {
      double total = 0;
      string side = "left";
      for (int j = 0; j < 2; j++) {
        for (int i = 0; i < PANEL_GRID.Rows.Count; i++) {
          if (i + 2 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 2].Cells[$"breaker_{side}"].Value as string != "3") {
            if (i + 1 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 1].Cells[$"breaker_{side}"].Value as string != "3") {
              if (i < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i].Cells[$"breaker_{side}"].Value as string != "3") {
                total += SafeConvertToDouble(PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value?.ToString());
                total += SafeConvertToDouble(PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value?.ToString());
                total += SafeConvertToDouble(PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value?.ToString());
              }
            }
          }
        }
        side = "right";
      }
      return total;
    }

    private double AggregateThreePhaseLoads() {
      double total = 0;
      string side = "left";
      for (int j = 0; j < 2; j++) {
        for (int i = 0; i < PANEL_GRID.Rows.Count; i++) {
          if (i + 2 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 2].Cells[$"breaker_{side}"].Value as string == "3") {
            total += SafeConvertToDouble(PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value?.ToString());
            total += SafeConvertToDouble(PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value?.ToString());
            total += SafeConvertToDouble(PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value?.ToString());
          }
          else if (i + 1 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 1].Cells[$"breaker_{side}"].Value as string == "3") {
            total += SafeConvertToDouble(PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value?.ToString());
            total += SafeConvertToDouble(PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value?.ToString());
            total += SafeConvertToDouble(PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value?.ToString());
          }
          else if (i < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i].Cells[$"breaker_{side}"].Value as string == "3") {
            total += SafeConvertToDouble(PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value?.ToString());
            total += SafeConvertToDouble(PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value?.ToString());
            total += SafeConvertToDouble(PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value?.ToString());
          }
        }
        side = "right";
      }
      return total;
    }
    private void PHASE_SUM_GRID_CellValueChanged(object sender, DataGridViewCellEventArgs e) {
      if (e.RowIndex == 0 && (e.ColumnIndex == 0 || e.ColumnIndex == 1 || e.ColumnIndex == 2)) {
        UpdatePerCellValueChange();
      }
    }

    private void ColorFeederAmpsBox(double feederAmps, double busRating) {
      if (feederAmps > busRating) {
        // turn bg red
        foreach (DataGridViewRow row in FEEDER_AMP_GRID.Rows) {
          foreach (DataGridViewCell cell in row.Cells) {
            cell.Style.SelectionBackColor = Color.Crimson;
            cell.Style.SelectionForeColor = Color.White;
          }
        }
      }
      else if (feederAmps > 0.8 * busRating) {
        // turn bg yellow
        foreach (DataGridViewRow row in FEEDER_AMP_GRID.Rows) {
          foreach (DataGridViewCell cell in row.Cells) {
            cell.Style.SelectionBackColor = Color.DarkGoldenrod;
            cell.Style.SelectionForeColor = Color.White;
          }
        }
      }
      else {
        // turn bg green
        foreach (DataGridViewRow row in FEEDER_AMP_GRID.Rows) {
          foreach (DataGridViewCell cell in row.Cells) {
            cell.Style.SelectionBackColor = Color.DarkCyan;
            cell.Style.SelectionForeColor = Color.White;
          }
        }
      }
    }

    private void SetPanelLoadLispVars(double totalKva, double feederAmps) {
      Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      doc.SetLispSymbol($"panel_{Name}_kva", Math.Round(totalKva, 1) + " KVA");
      doc.SetLispSymbol($"panel_{Name}_a", feederAmps.ToString() + " A");
    }

    public void UpdatePerCellValueChange() {
      if (isLoading) {
        return;
      }
      object phaseVoltageObj = PHASE_VOLTAGE_COMBOBOX.SelectedItem;
      object lineVoltageObj = LINE_VOLTAGE_COMBOBOX.SelectedItem;
      double phaseVoltage;
      double lineVoltage;
      if (phaseVoltageObj != null) {
        phaseVoltage = Convert.ToDouble(phaseVoltageObj);
        lineVoltage = Convert.ToDouble(lineVoltageObj);
      }
      else {
        return;
      }
      double feederAmps = 0;
      double busRating = 0;
      double totalKva = 0;
      double sum = 0;
      double safetyFactor = 1.0;
      if (!String.IsNullOrEmpty(BUS_RATING_INPUT.Text)) {
        busRating = Convert.ToDouble(BUS_RATING_INPUT.Text);
      }
      if (SAFETY_FACTOR_CHECKBOX.Enabled && SAFETY_FACTOR_CHECKBOX.Checked && Regex.IsMatch(SAFETY_FACTOR_TEXTBOX.Text, @"^\d*\.?\d*$")) {
        safetyFactor = Convert.ToDouble(SAFETY_FACTOR_TEXTBOX.Text);
      }
      if (lineVoltage == 240 && phaseVoltage == 120 && this.is3Ph) {
        // perform high leg calculation
        double singlePhaseLoads = AggregateSinglePhaseLoads();
        double threePhaseLoads = AggregateThreePhaseLoads();
        sum = singlePhaseLoads + threePhaseLoads;
        double singlePhaseAmperage = singlePhaseLoads / 240;
        double threePhaseAmperage = threePhaseLoads / 240 / 1.732;
        feederAmps = singlePhaseAmperage + threePhaseAmperage;
        FEEDER_AMP_GRID.Rows[0].Cells[0].Value = Math.Round(feederAmps, 1);
        totalKva = CalculatePanelLoad(sum) * safetyFactor;
        PANEL_LOAD_GRID.Rows[0].Cells[0].Value = Math.Round(totalKva, 1);
        ColorFeederAmpsBox(feederAmps, busRating);
        SetPanelLoadLispVars(totalKva, feederAmps);
        return;
      }
      double phA = Convert.ToDouble(PHASE_SUM_GRID.Rows[0].Cells[0].Value ?? 0);
      double phB = Convert.ToDouble(PHASE_SUM_GRID.Rows[0].Cells[1].Value ?? 0);
      double phC = 0;
      sum = phA + phB;
      int poles = 2;
      if (PHASE_SUM_GRID.ColumnCount > 2) {
        phC = Convert.ToDouble(PHASE_SUM_GRID.Rows[0].Cells[2].Value ?? 0);
        sum += phC;
        poles = 3;
      }
      mainForm.UpdateLclLml();
      
      TOTAL_VA_GRID.Rows[0].Cells[0].Value = CalculateTotalVA(sum);

      // Handle LCL
      if (!string.IsNullOrEmpty(LCL.Text) && LCL.Text != "0") {
        sum += Math.Round(Convert.ToDouble(LCL.Text) * 0.25, 0);
      }

      // Handle LML
      if (!string.IsNullOrEmpty(LML.Text) && LML.Text != "0") {
        sum += Math.Round(Convert.ToDouble(LML.Text) * 0.25, 0);
      }
      totalKva = CalculatePanelLoad(sum) * safetyFactor;
      PANEL_LOAD_GRID.Rows[0].Cells[0].Value = Math.Round(totalKva, 1);
      if (phaseVoltageObj != null) {
        double yFactor = 1;
        if ((lineVoltage == 208.0 || lineVoltage == 480.0) && this.is3Ph) {
          yFactor = 1.732;
        }
        if (phaseVoltage != 0) {
          
          if (LCL.Text == "0" && LML.Text == "0") {
            if (DISTRIBUTION_SECTION_CHECKBOX.Checked) {
              feederAmps = Math.Round(totalKva * 1000 / lineVoltage / yFactor, 1);
              FEEDER_AMP_GRID.Rows[0].Cells[0].Value = feederAmps;
            }
            else {
              feederAmps = CalculateFeederAmps(phA, phB, phC, phaseVoltage) * safetyFactor;
              FEEDER_AMP_GRID.Rows[0].Cells[0].Value = feederAmps;
            }
          }
          else {
            feederAmps = Math.Round(sum / (phaseVoltage * poles), 1);
            FEEDER_AMP_GRID.Rows[0].Cells[0].Value = feederAmps;
          }
          ColorFeederAmpsBox(feederAmps, busRating);
          SetPanelLoadLispVars(totalKva, feederAmps);
        }
      }
    }
    public double CalculateFeederAmps(double phA, double phB, double phC, double lineVoltage) {
      if (lineVoltage == 0) {
        return 0;
      }

      double maxVal = Math.Max(Math.Max(phA, phB), phC);
      
      return Math.Round(maxVal / lineVoltage, 1);
    }

    public void UpdateLclLmlLabels(int lcl, int lml) {
      if (!LCL_OVERRIDE.Checked) {
        LCL.Text = $"{lcl}";
      }
      if (!LML_OVERRIDE.Checked) {
        LML.Text = $"{lml}";
      }
    }

    private void SaveLCLLMLObjectAsJson(object LCLLMLObject) {
      string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
      string filePath = Path.Combine(desktopPath, "LCLLMLObject.json");

      try {
        string jsonString = JsonConvert.SerializeObject(LCLLMLObject, Formatting.Indented);
        File.WriteAllText(filePath, jsonString);
        Console.WriteLine("LCLLMLObject saved successfully.");
      }
      catch (Exception ex) {
        Console.WriteLine($"Error saving LCLLMLObject: {ex.Message}");
      }
    }

    private void ADD_ROW_BUTTON_Click(object sender, EventArgs e) {
      PANEL_GRID.Rows.Add();

      if (PANEL_GRID.Rows.Count > 21) {
        PANEL_GRID.Width = 1047 + 15;
      }
    }

    private void CREATE_PANEL_BUTTON_Click(object sender, EventArgs e) {
      Dictionary<string, object> panelDataList = RetrieveDataFromModal();

      using (
        DocumentLock docLock =
          Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
      ) {
        this.mainForm.Close();
        myCommandsInstance.CreatePanel(panelDataList);

        Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.WindowState = Autodesk
          .AutoCAD
          .Windows
          .Window
          .State
          .Maximized;
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Window.Focus();
      }
    }

    private void CREATE_LOAD_SUMMARY_BUTTON_Click(object sender, EventArgs e) {
      this.mainForm.SavePanelDataToLocalJsonFile();
      Dictionary<string, object> panelDataList = RetrieveDataFromModal();
      List<Dictionary<string, object>> panelStorage = this.mainForm.RetrieveSavedPanelData();
      using (
        DocumentLock docLock =
          Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument()
      ) {
        this.mainForm.Close();
        myCommandsInstance.CreateLoadSummary(panelDataList, panelStorage);

        Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.WindowState = Autodesk
          .AutoCAD
          .Windows
          .Window
          .State
          .Maximized;
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Window.Focus();
      }
    }

     private void DELETE_ROW_BUTTON_Click(object sender, EventArgs e) {
      if (PANEL_GRID.Rows.Count > 0) {
        var lastRow = PANEL_GRID.Rows[PANEL_GRID.Rows.Count - 1];
        var phaseCells = new List<string>
        {
          "phase_a_left",
          "phase_b_left",
          "phase_a_right",
          "phase_b_right"
        };

        if (PHASE_SUM_GRID.ColumnCount > 2) {
          phaseCells.Add("phase_c_left");
          phaseCells.Add("phase_c_right");
        }

        foreach (var cell in phaseCells) {
          lastRow.Cells[cell].Value = "0";
        }

        PANEL_GRID.Rows.RemoveAt(PANEL_GRID.Rows.Count - 1);

        if (PANEL_GRID.Rows.Count <= 21) {
          PANEL_GRID.Width = 1047;
        }
      }
    }

    private void DELETE_PANEL_BUTTON_Click(object sender, EventArgs e) {
      this.mainForm.DeletePanel(this);
    }

    private void INFO_LABEL_CLICK(object sender, EventArgs e) {
    }

    private void NOTES_BUTTON_Click(object sender, EventArgs e) {
      if (this.notesForm == null || this.notesForm.IsDisposed) {
        this.notesForm = new NoteForm(this);
        this.notesForm.Show();
        this.notesForm.Text = $"Panel '{PANEL_NAME_INPUT.Text}' Notes";
      }
      else {
        if (!this.notesForm.Visible) {
          this.notesForm.Show();
        }
        this.notesForm.BringToFront();
      }
    }

    private void APPLY_BUTTON_Click(object sender, EventArgs e) {
      string selectedValue = APPLY_COMBOBOX.SelectedItem.ToString();

      List<string> columnNames = new List<string> { "description" };

      foreach (DataGridViewCell cell in PANEL_GRID.SelectedCells) {
        if (columnNames.Any(cell.OwningColumn.Name.Contains)) {
          if (cell.Tag == null) {
            cell.Tag = selectedValue;
          }
          else {
            cell.Tag = $"{cell.Tag}|{selectedValue}";
          }
          cell.Style.BackColor = Color.Yellow;
        }
      }

      CalculateBreakerLoad();
      UpdatePerCellValueChange();
    }

    private void APPLY_COMBOBOX_SelectedIndexChanged(object sender, EventArgs e) {
      UpdateCellBackgroundColor();
    }

    private void STATUS_COMBOBOX_SelectedIndexChanged(object sender, EventArgs e) {
      var default_existing_message =
        "DENOTES EXISTING CIRCUIT BREAKER TO REMAIN; ALL OTHERS ARE NEW.";
      var default_new_message = "65 KAIC SERIES RATED OR MATCH FAULT CURRENT AT SITE.";

      if (STATUS_COMBOBOX.SelectedItem != null) {
        if (
          STATUS_COMBOBOX.SelectedItem.ToString() == "EXISTING"
          || STATUS_COMBOBOX.SelectedItem.ToString() == "RELOCATED"
        ) {
          if (!this.notesStorage.Contains(default_existing_message)) {
            this.notesStorage.Add(default_existing_message);
          }
          if (this.notesStorage.Contains(default_new_message)) {
            this.notesStorage.Remove(default_new_message);
          }
          RemoveTagsFromCells(default_new_message);
        }
        else {
          if (!this.notesStorage.Contains(default_new_message)) {
            this.notesStorage.Add(default_new_message);
          }
          if (this.notesStorage.Contains(default_existing_message)) {
            this.notesStorage.Remove(default_existing_message);
          }
          RemoveTagsFromCells(default_existing_message);
        }
        UpdateApplyComboboxToMatchStorage();
      }
    }

    private void REMOVE_NOTE_BUTTON_Click(object sender, EventArgs e) {
      string selectedValue = APPLY_COMBOBOX.SelectedItem.ToString();
      List<string> columnNames = new List<string> { "description" };

      foreach (DataGridViewCell cell in PANEL_GRID.SelectedCells) {
        if (columnNames.Any(cell.OwningColumn.Name.Contains)) {
          if (cell.Tag != null && cell.Tag.ToString().Contains(selectedValue)) {
            cell.Tag = cell.Tag.ToString().Replace(selectedValue, "").Trim('|');
            cell.Style.BackColor = Color.Empty;
          }
        }
      }

      CalculateBreakerLoad();
      UpdatePerCellValueChange();
    }

    private void REPLACE_BUTTON_Click(object sender, EventArgs e) {
      string findText = FIND_BOX.Text;
      string replaceText = REPLACE_BOX.Text;

      if (string.IsNullOrEmpty(findText)) {
        MessageBox.Show("Please enter a text to find.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
      }

      foreach (DataGridViewCell cell in PANEL_GRID.SelectedCells) {
        if (cell.Value != null && cell.Value.ToString().Contains(findText)) {
          cell.Value = cell.Value.ToString().Replace(findText, replaceText);
        }
      }
    }

    public string GetNewOrExisting() {
      return STATUS_COMBOBOX.Text;
    }

    public void AddListeners() {
      PANEL_GRID.KeyDown += new KeyEventHandler(this.PANEL_GRID_KeyDown);
      PANEL_GRID.CellBeginEdit += new DataGridViewCellCancelEventHandler(
        this.PANEL_GRID_CellBeginEdit
      );
      PANEL_GRID.CellValueChanged += new DataGridViewCellEventHandler(
        this.PANEL_GRID_CellValueChanged
      );
      PHASE_SUM_GRID.CellValueChanged += new DataGridViewCellEventHandler(
        this.PHASE_SUM_GRID_CellValueChanged
      );
      PANEL_GRID.CellFormatting += PANEL_GRID_CellFormatting;
      PANEL_GRID.CellClick += new DataGridViewCellEventHandler(this.PANEL_GRID_CellClick);
      PANEL_NAME_INPUT.Click += (sender, e) => {
        PANEL_NAME_INPUT.SelectAll();
      };
      PANEL_LOCATION_INPUT.Click += (sender, e) => {
        PANEL_LOCATION_INPUT.SelectAll();
      };
      PANEL_GRID.CellPainting += new DataGridViewCellPaintingEventHandler(PANEL_GRID_CellPainting);
      MAIN_INPUT.Click += (sender, e) => {
        MAIN_INPUT.SelectAll();
      };
      BUS_RATING_INPUT.Click += (sender, e) => {
        BUS_RATING_INPUT.SelectAll();
      };
      EXISTING_BUTTON.Click += new System.EventHandler(this.EXISTING_BUTTON_Click);
      RELOCATE_BUTTON.Click += new System.EventHandler(this.RELOCATE_BUTTON_Click);
      LCL.TextChanged += new EventHandler(LCL_LML_TextChanged);
      LML.TextChanged += new EventHandler(LCL_LML_TextChanged);
    }

    private void LCL_LML_TextChanged(object sender, EventArgs e) {
      if (LCL_OVERRIDE.Checked || LML_OVERRIDE.Checked) {
        UpdatePerCellValueChange();
      }
    }

    private void LCL_OVERRIDE_CheckedChanged(object sender, EventArgs e) {
      if (LCL.Enabled) {
        LCL.Enabled = false;
        UpdatePerCellValueChange();
      }
      else {
        LCL.Enabled = true;
      }
    }

    private void LML_OVERRIDE_CheckedChanged(object sender, EventArgs e) {
      if (LML.Enabled) {
        LML.Enabled = false;
        UpdatePerCellValueChange();
      }
      else {
        LML.Enabled = true;
      }
    }

    private void SAFETY_FACTOR_CheckChanged(object sender, EventArgs e) {
      UpdatePerCellValueChange();
    }

    private void VoltageCombobox_SelectedValueChanged(object sender, EventArgs e) {
      UpdatePerCellValueChange();
      if (is3Ph) {
        Color3pPanel(sender, e);
      }
      SetWarnings();
    }

    private void ADD_ALL_PANELS_BUTTON_Click(object sender, EventArgs e) {
      this.mainForm.SavePanelDataToLocalJsonFile();
      // iterate over all panels
      List<Dictionary<string, object>> allPanelData = this.mainForm.RetrieveSavedPanelData();
      int index = 0;
      int prevIndex = 0;
      int nextIndex = 0;
      string side = "left";
      int phase = Convert.ToInt32(PHASE_COMBOBOX.Text);
      if (phase == 1) {
        phase = 2;
      }
      void add(string panelName) {
        PANEL_GRID.Rows[index].Cells[$"description_{side}"].Value = "PANEL " + panelName;
        PANEL_GRID.Rows[index + 1].Cells[$"description_{side}"].Value = "";
        if (phase == 3) {
          PANEL_GRID.Rows[index + 2].Cells[$"description_{side}"].Value = "";
        }
        index += phase;
        if (index >= PANEL_GRID.Rows.Count && side == "left") {
          side = "right";
          nextIndex = index;
          index = prevIndex;
          prevIndex = nextIndex;
        }
        else if (index >= PANEL_GRID.Rows.Count && side == "right") {
          side = "left";
          for (int i = 0; i < phase; i++) {
            ADD_ROW_BUTTON_Click(sender, e);
          }
        }
      }
      foreach (Dictionary<string, object> panel in allPanelData) {
        if (panel.TryGetValue("distribution_section", out object value)) {
          if (value is bool boolValue && boolValue == false) {
            if (panel.TryGetValue("panel", out object panelName)) {
              panelName = panelName.ToString().Replace("'", "");
              if (panel.TryGetValue("fed_from", out object fedFrom)) {
                fedFrom = fedFrom.ToString();
                if (!fedFrom.ToString().StartsWith("PANEL") && !fedFrom.ToString().StartsWith("SUBPANEL")) {
                  add(panelName as string);
                }
              }
              else {
                add(panelName as string);
              }
            }
          }
        }
      }
      //GetSubPanels();
      if (GetPanelLoad() > 0) {
        ADD_ALL_PANELS_BUTTON.Enabled = false;
      }
    }

    public double GetLclOverride() {
      if (LCL_OVERRIDE.Checked) {
        double result;
        if (double.TryParse(LCL.Text, out result)) {
          return result;
        }
        else {
          return 0;
        }
      }
      else {
        return 0;
      }
    }

    public double GetLmlOverride() {
      if (LML_OVERRIDE.Checked) {
        double result;
        if (double.TryParse(LML.Text, out result)) {
          return result;
        }
        else {
          return 0;
        }
      }
      else {
        return 0;
      }
    }

    private string ConvertAtoVA(string aValue, int numPoles, string voltage) {
      string sanitized = Regex.Replace(aValue, @" *A *", "");
      if (numPoles == 3) {
        return (Math.Round(Convert.ToDouble(sanitized) * Convert.ToDouble(voltage) / 1.732, 0)).ToString();
      }
      else if (numPoles == 2) {
        return (Math.Round(Convert.ToDouble(sanitized) * Convert.ToDouble(voltage) / 2, 0)).ToString();
      }
      else{
        return (Math.Round(Convert.ToDouble(sanitized) * Convert.ToDouble(voltage), 0)).ToString();
      }
    }
    private string ConvertHpToVa(string hpValue, int numPhases, string voltage) {
      string sanitized = Regex.Replace(hpValue, @" +HP", "HP");
      sanitized = Regex.Replace(sanitized, @"\p{Zs}+", "+");
      sanitized = Regex.Replace(sanitized, @"-+", "+");
      sanitized = sanitized.Replace("HP", "");
      if (!Regex.IsMatch(sanitized, @"^\d+((\+\d)?\/\d)?$")) {
        return "-1";
      }
      System.Data.DataTable dt = new System.Data.DataTable();
      double sumObject = Convert.ToDouble(dt.Compute(sanitized, null));
      string phaseVA = "";

      if (numPhases == 1) {
        switch (sumObject) {
          case var _ when sumObject > 7.5: { // 10
              if (voltage == "120") {
                phaseVA = "12000";
              }
              if (voltage == "208") {
                phaseVA = "5720";
              }
              if (voltage == "240") {
                phaseVA = "6000";
              }
              break;
            };
          case var _ when sumObject > 5: { // 7.5
              if (voltage == "120") {
                phaseVA = "9600";
              }
              if (voltage == "208") {
                phaseVA = "4576";
              }
              if (voltage == "240") {
                phaseVA = "4800";
              }
              break;
            };
          case var _ when sumObject > 3: { // 5
              if (voltage == "120") {
                phaseVA = "6720";
              }
              if (voltage == "208") {
                phaseVA = "3220";
              }
              if (voltage == "240") {
                phaseVA = "3360";
              }
              break;
            };
          case var _ when sumObject > 2: { // 3
              if (voltage == "120") {
                phaseVA = "4080";
              }
              if (voltage == "208") {
                phaseVA = "1945";
              }
              if (voltage == "240") {
                phaseVA = "2040";
              }
              break;
            };
          case var _ when sumObject > 1.5: { // 2
              if (voltage == "120") {
                phaseVA = "2880";
              }
              if (voltage == "208") {
                phaseVA = "1373";
              }
              if (voltage == "240") {
                phaseVA = "2040";
              }
              break;
            };
          case var _ when sumObject > 1: { // 1 1/2
              if (voltage == "120") {
                phaseVA = "2400";
              }
              if (voltage == "208") {
                phaseVA = "1144";
              }
              if (voltage == "240") {
                phaseVA = "1200";
              }
              break;
            };
          case var _ when sumObject > 0.75: { // 1
              if (voltage == "120") {
                phaseVA = "1920";
              }
              if (voltage == "208") {
                phaseVA = "915";
              }
              if (voltage == "240") {
                phaseVA = "960";
              }
              break;
            };
          case var _ when sumObject > 0.5: { // 3/4
              if (voltage == "120") {
                phaseVA = "1656";
              }
              if (voltage == "208") {
                phaseVA = "791";
              }
              if (voltage == "240") {
                phaseVA = "828";
              }
              break;
            };
          case var _ when sumObject > 0.34: { // 1/2
              if (voltage == "120") {
                phaseVA = "1176";
              }
              if (voltage == "208") {
                phaseVA = "562";
              }
              if (voltage == "240") {
                phaseVA = "588";
              }
              break;
            };
          case var _ when sumObject > 0.25: { // 1/3
              if (voltage == "120") {
                phaseVA = "864";
              }
              if (voltage == "208") {
                phaseVA = "416";
              }
              if (voltage == "240") {
                phaseVA = "432";
              }
              break;
            };
          case var _ when sumObject > 0.167: { // 1/4
              if (voltage == "120") {
                phaseVA = "696";
              }
              if (voltage == "208") {
                phaseVA = "333";
              }
              if (voltage == "240") {
                phaseVA = "348";
              }
              break;
            };
          case var _ when sumObject <= 0.167: { // 1/6
              if (voltage == "120") {
                phaseVA = "528";
              }
              if (voltage == "208") {
                phaseVA = "250";
              }
              if (voltage == "240") {
                phaseVA = "264";
              }
              break;
            };
        }
      }

      if (numPhases == 3) {
        switch (sumObject) {
          case var _ when sumObject > 15: { // 20
              if (voltage == "208") {
                phaseVA = Math.Round(59.4 * 208.0).ToString();
              }
              if (voltage == "240") {
                phaseVA = Math.Round(54 * 230.0).ToString();
              }
              if (voltage == "480") {
                phaseVA = Math.Round(27 * 460.0).ToString();
              }
              break;
            }
          case var _ when sumObject > 10: { // 15
              if (voltage == "208") {
                phaseVA = Math.Round(46.2 * 208.0).ToString();
              }
              if (voltage == "240") {
                phaseVA = Math.Round(42 * 230.0).ToString();
              }
              if (voltage == "480") {
                phaseVA = Math.Round(21.0 * 460.0).ToString();
              }
              break;
            };
          case var _ when sumObject > 7.5: { // 10
              if (voltage == "208") {
                phaseVA = Math.Round(30.8 * 208.0).ToString();
              }
              if (voltage == "240") {
                phaseVA = Math.Round(28 * 230.0).ToString();
              }
              if (voltage == "480") {
                phaseVA = Math.Round(14.0 * 460.0).ToString();
              }
              break;
            };
          case var _ when sumObject > 5: { // 7 1/2
              if (voltage == "208") {
                phaseVA = Math.Round(24.2 * 208.0).ToString();
              }
              if (voltage == "240") {
                phaseVA = Math.Round(22 * 230.0).ToString();
              }
              if (voltage == "480") {
                phaseVA = Math.Round(11.0 * 460.0).ToString();
              }
              break;
            };
          case var _ when sumObject > 3: { // 5
              if (voltage == "208") {
                phaseVA = Math.Round(16.7 * 208.0).ToString();
              }
              if (voltage == "240") {
                phaseVA = Math.Round(15.2 * 230.0).ToString();
              }
              if (voltage == "480") {
                phaseVA = Math.Round(7.6 * 460.0).ToString();
              }
              break;
            };
          case var _ when sumObject > 2: { // 3
              if (voltage == "208") {
                phaseVA = Math.Round(10.6 * 208.0).ToString();
              }
              if (voltage == "240") {
                phaseVA = Math.Round(9.6 * 230.0).ToString();
              }
              if (voltage == "480") {
                phaseVA = Math.Round(4.8 * 460.0).ToString();
              }
              break;
            };
          case var _ when sumObject > 1.5: { // 2
              if (voltage == "208") {
                phaseVA = Math.Round(7.5 * 208.0).ToString();
              }
              if (voltage == "240") {
                phaseVA = Math.Round(6.8 * 230.0).ToString();
              }
              if (voltage == "480") {
                phaseVA = Math.Round(3.4 * 460.0).ToString();
              }
              break;
            };
          case var _ when sumObject > 1: { // 1 1/2
              if (voltage == "208") {
                phaseVA = Math.Round(6.6 * 208.0).ToString();
              }
              if (voltage == "240") {
                phaseVA = Math.Round(6 * 230.0).ToString();
              }
              if (voltage == "480") {
                phaseVA = Math.Round(3.0 * 460.0).ToString();
              }
              break;
            };
          case var _ when sumObject > 0.75: { // 1
              if (voltage == "208") {
                phaseVA = Math.Round(4.6 * 208.0).ToString();
              }
              if (voltage == "240") {
                phaseVA = Math.Round(4.2 * 230.0).ToString();
              }
              if (voltage == "480") {
                phaseVA = Math.Round(2.1 * 460.0).ToString();
              }
              break;
            };
          case var _ when sumObject > 0.5: { // 3/4
              if (voltage == "208") {
                phaseVA = Math.Round(3.5 * 208.0).ToString();
              }
              if (voltage == "240") {
                phaseVA = Math.Round(3.2 * 230.0).ToString();
              }
              if (voltage == "480") {
                phaseVA = Math.Round(1.6 * 460.0).ToString();
              }
              break;
            };
          case var _ when sumObject <= 0.5: { // 1/2
              if (voltage == "208") {
                phaseVA = Math.Round(2.4 * 208.0).ToString();
              }
              if (voltage == "240") {
                phaseVA = Math.Round(2.2 * 230.0).ToString();
              }
              if (voltage == "480") {
                phaseVA = Math.Round(1.1 * 460.0).ToString();
              }
              break;
            };
        }
      }
      return phaseVA;
    }

    private void ConvertHpToVaBySide3Ph(string side) {
      int i = 0;
      while (i < PANEL_GRID.Rows.Count) {
        string phaseA = PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value?.ToString().ToUpper().Replace("\r", "");
        string phaseB = PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value?.ToString().ToUpper().Replace("\r", "");
        string phaseC = PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value?.ToString().ToUpper().Replace("\r", "");

        if (!String.IsNullOrEmpty(phaseA) && phaseA.EndsWith("HP")) {
          if (i + 1 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 1].Cells[$"breaker_{side}"].Value as string == "2") {
            phaseA = ConvertHpToVa(phaseA, 1, LINE_VOLTAGE_COMBOBOX.SelectedItem as string);
            phaseB = phaseA;
            if (phaseA == "-1" || phaseB == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side} and phase_b_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value = phaseA;
            PANEL_GRID.Rows[i + 1].Cells[$"phase_b_{side}"].Value = phaseB;
            i += 2;
          }
          else if (i + 2 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 2].Cells[$"breaker_{side}"].Value as string == "3") {
            phaseA = ConvertHpToVa(phaseA, 3, LINE_VOLTAGE_COMBOBOX.SelectedItem as string);
            phaseB = phaseA;
            phaseC = phaseA;
            if (phaseA == "-1" || phaseB == "-1" || phaseC == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side}, phase_b_{side}, phase_c_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value = phaseA;
            PANEL_GRID.Rows[i + 1].Cells[$"phase_b_{side}"].Value = phaseB;
            PANEL_GRID.Rows[i + 2].Cells[$"phase_c_{side}"].Value = phaseC;
            i += 3;
          }
          else {
            phaseA = ConvertHpToVa(phaseA, 1, PHASE_VOLTAGE_COMBOBOX.SelectedItem as string);
            if (phaseA == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value = phaseA;
            i++;
          }
        }
        else if (!String.IsNullOrEmpty(phaseB) && phaseB.EndsWith("HP")) {
          if (i + 1 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 1].Cells[$"breaker_{side}"].Value as string == "2") {
            phaseB = ConvertHpToVa(phaseB, 1, LINE_VOLTAGE_COMBOBOX.SelectedItem as string);
            phaseC = phaseB;
            if (phaseB == "-1" || phaseC == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side} and phase_b_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value = phaseB;
            PANEL_GRID.Rows[i + 1].Cells[$"phase_c_{side}"].Value = phaseC;
            i += 2;
          }
          else if (i + 2 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 2].Cells[$"breaker_{side}"].Value as string == "3") {
            phaseB = ConvertHpToVa(phaseB, 3, LINE_VOLTAGE_COMBOBOX.SelectedItem as string);
            phaseC = phaseB;
            phaseA = phaseB;
            if (phaseB == "-1" || phaseC == "-1" || phaseA == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side}, phase_b_{side}, phase_c_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value = phaseB;
            PANEL_GRID.Rows[i + 1].Cells[$"phase_c_{side}"].Value = phaseC;
            PANEL_GRID.Rows[i + 2].Cells[$"phase_a_{side}"].Value = phaseA;
            i += 3;
          }
          else {
            phaseB = ConvertHpToVa(phaseB, 1, PHASE_VOLTAGE_COMBOBOX.SelectedItem as string);
            if (phaseB == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value = phaseB;
            i++;
          }
        }
        else if (!String.IsNullOrEmpty(phaseC) && phaseC.EndsWith("HP")) {
          if (i + 1 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 1].Cells[$"breaker_{side}"].Value as string == "2") {
            phaseC = ConvertHpToVa(phaseC, 1, LINE_VOLTAGE_COMBOBOX.SelectedItem as string);
            phaseA = phaseC;
            if (phaseC == "-1" || phaseA == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side} and phase_b_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value = phaseC;
            PANEL_GRID.Rows[i + 1].Cells[$"phase_a_{side}"].Value = phaseA;
            i += 2;
          }
          else if (i + 2 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 2].Cells[$"breaker_{side}"].Value as string == "3") {
            phaseC = ConvertHpToVa(phaseC, 3, LINE_VOLTAGE_COMBOBOX.SelectedItem as string);
            phaseA = phaseC;
            phaseB = phaseC;
            if (phaseC == "-1" || phaseA == "-1" || phaseB == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side}, phase_b_{side}, phase_c_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value = phaseC;
            PANEL_GRID.Rows[i + 1].Cells[$"phase_a_{side}"].Value = phaseA;
            PANEL_GRID.Rows[i + 2].Cells[$"phase_b_{side}"].Value = phaseB;
            i += 3;
          }
          else {
            phaseC = ConvertHpToVa(phaseC, 1, PHASE_VOLTAGE_COMBOBOX.SelectedItem as string);
            if (phaseC == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value = phaseC;
            i++;
          }
        }
        else {
          i += 1;
        }
      }

    }

    private void ConvertHpToVaBySide1Ph(string side) {
      int i = 0;
      while (i < PANEL_GRID.Rows.Count) {
        string phaseA = PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value?.ToString().ToUpper().Replace("\r", "");
        string phaseB = PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value?.ToString().ToUpper().Replace("\r", "");

        if (!String.IsNullOrEmpty(phaseA) && phaseA.EndsWith("HP")) {
          if (i + 1 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 1].Cells[$"breaker_{side}"].Value as string == "2") {
            phaseA = ConvertHpToVa(phaseA, 1, LINE_VOLTAGE_COMBOBOX.SelectedItem as string);
            phaseB = phaseA;
            if (phaseA == "-1" || phaseB == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side} and phase_b_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value = phaseA;
            PANEL_GRID.Rows[i + 1].Cells[$"phase_b_{side}"].Value = phaseB;
            i += 2;
          }
          else {
            phaseA = ConvertHpToVa(phaseA, 1, PHASE_VOLTAGE_COMBOBOX.SelectedItem as string);
            if (phaseA == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value = phaseA;
            i++;
          }
        }
        else if (!String.IsNullOrEmpty(phaseB) && phaseB.EndsWith("HP")) {
          if (i + 1 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 1].Cells[$"breaker_{side}"].Value as string == "2") {
            phaseB = ConvertHpToVa(phaseB, 1, LINE_VOLTAGE_COMBOBOX.SelectedItem as string);
            phaseA = phaseB;
            if (phaseA == "-1" || phaseB == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side} and phase_b_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value = phaseB;
            PANEL_GRID.Rows[i + 1].Cells[$"phase_a_{side}"].Value = phaseA;
            i += 2;
          }
          else {
            phaseB = ConvertHpToVa(phaseB, 1, PHASE_VOLTAGE_COMBOBOX.SelectedItem as string);
            if (phaseB == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value = phaseB;
            i++;
          }
        }
        else {
          i += 1;
        }
      }

    }

    private void ConvertHpToVa_Click(object sender, EventArgs e) {
      if (PANEL_GRID.Columns["phase_c_left"] != null) {
        ConvertHpToVaBySide3Ph("left");
        ConvertHpToVaBySide3Ph("right");
      }
      else {
        ConvertHpToVaBySide1Ph("left");
        ConvertHpToVaBySide1Ph("right");
      }
    }

    public bool IsUuid(string str) {
      return Regex.IsMatch(str, @"^[a-fA-F0-9]{8}(-[a-fA-F0-9]{4}){3}-[a-fA-F0-9]{12}$");
    }

    internal void UpdateSubpanelName(string oldSubpanelName, string newSubpanelName) {
      for (int i = 0; i < PANEL_GRID.Rows.Count; i++) {
        if ((PANEL_GRID.Rows[i].Cells["description_left"].Value as string).ToUpper() == oldSubpanelName.ToUpper()) {
          PANEL_GRID.Rows[i].Cells["description_left"].Value = newSubpanelName.ToUpper();
        }
        if ((PANEL_GRID.Rows[i].Cells["description_right"].Value as string).ToUpper() == oldSubpanelName.ToUpper()) {
          PANEL_GRID.Rows[i].Cells["description_right"].Value = newSubpanelName.ToUpper();
        }
      }
    }

    internal void UpdateSubpanelFedFrom() {
      string side = "left";
      for (int j = 0; j < 2; j++) {
        for (int i = 0; i < PANEL_GRID.Rows.Count; i++) {
          if ((PANEL_GRID.Rows[i].Cells[$"description_{side}"].Value as string).ToUpper().StartsWith("PANEL")) {
            PanelUserControl panelControl = (PanelUserControl)this.mainForm.FindUserControl((PANEL_GRID.Rows[i].Cells[$"description_{side}"].Value as string).ToUpper());
            if (panelControl != null) {
              string name = this.Name;
              TextBox fedFromTextbox =
                  panelControl.Controls.Find("FED_FROM_TEXTBOX", true).FirstOrDefault() as TextBox;
              Label phaseWarningLabel = panelControl.Controls.Find("PHASE_WARNING_LABEL", true).FirstOrDefault() as Label;
              ComboBox phaseComboBox = panelControl.Controls.Find("PHASE_COMBOBOX", true).FirstOrDefault() as ComboBox;
              ComboBox wireComboBox = panelControl.Controls.Find("WIRE_COMBOBOX", true).FirstOrDefault() as ComboBox;
              if (DISTRIBUTION_SECTION_CHECKBOX.Checked) {
                fedFromTextbox.Text = name;
              }
              else {
                fedFromTextbox.Text = "PANEL " + name;
              }
              phaseWarningLabel.Visible = true;
              phaseComboBox.Enabled = false;
              wireComboBox.Enabled = false;
            }
          }
        }
        side = "right";
      }
    }

    internal void RemoveFedFrom(string panelName, bool check = true) {
      if (check) {
        if (panelName.ToUpper().Replace("DISTRIB. ", "") == FED_FROM_TEXTBOX.Text) {
          FED_FROM_TEXTBOX.Text = "";
          if (!contains3PhEquip) {
            PHASE_COMBOBOX.Enabled = true;
            WIRE_COMBOBOX.Enabled = true;
            PHASE_WARNING_LABEL.Visible = false;
          }
        }
        if (panelName.ToUpper() == FED_FROM_TEXTBOX.Text) {
          FED_FROM_TEXTBOX.Text = "";
          if (!contains3PhEquip) {
            PHASE_COMBOBOX.Enabled = true;
            WIRE_COMBOBOX.Enabled = true;
            PHASE_WARNING_LABEL.Visible = false;
          }
        }
      }
      else {
        FED_FROM_TEXTBOX.Text = "";
        if (!contains3PhEquip) {
          PHASE_COMBOBOX.Enabled = true;
          WIRE_COMBOBOX.Enabled = true;
          PHASE_WARNING_LABEL.Visible = false;
        }
      }
    }

    internal void RemoveSubpanel(string subpanelName, bool subpanelIs3Ph) {
      string side = "left";
      for (int j = 0; j < 2; j++) {
        for (int i = 0; i < PANEL_GRID.Rows.Count; i++) {
          if ((PANEL_GRID.Rows[i].Cells[$"description_{side}"].Value as string).ToUpper().Replace("PANEL ", "") == subpanelName.ToUpper()) {
            if (subpanelIs3Ph) {
              PANEL_GRID.Rows[i + 2].Cells[$"description_{side}"].Value = "";
              PANEL_GRID.Rows[i + 2].Cells[$"phase_a_{side}"].Value = "";
              PANEL_GRID.Rows[i + 2].Cells[$"phase_b_{side}"].Value = "";
              PANEL_GRID.Rows[i + 2].Cells[$"phase_c_{side}"].Value = "";
            }
            PANEL_GRID.Rows[i].Cells[$"description_{side}"].Value = "SPARE";
            PANEL_GRID.Rows[i + 1].Cells[$"description_{side}"].Value = "";
            PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value = "";
            PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value = "";
            PANEL_GRID.Rows[i + 1].Cells[$"phase_a_{side}"].Value = "";
            PANEL_GRID.Rows[i + 1].Cells[$"phase_b_{side}"].Value = "";
          }
        }
        side = "right";
      }
    }

    public bool Is3Ph() {
      return is3Ph;
    }

    public string GetId() {
      return id;
    }

    private void ConvertAToVaBySide3Ph(string side) {
      int i = 0;
      while (i < PANEL_GRID.Rows.Count) {
        string phaseA = PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value?.ToString().ToUpper().Replace("\r", "");
        string phaseB = PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value?.ToString().ToUpper().Replace("\r", "");
        string phaseC = PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value?.ToString().ToUpper().Replace("\r", "");

        if (!String.IsNullOrEmpty(phaseA) && phaseA.EndsWith("A")) {
          if (i + 1 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 1].Cells[$"breaker_{side}"].Value as string == "2") {
            phaseA = ConvertAtoVA(phaseA, 2, LINE_VOLTAGE_COMBOBOX.SelectedItem as string);
            phaseB = phaseA;
            if (phaseA == "-1" || phaseB == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side} and phase_b_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value = phaseA;
            PANEL_GRID.Rows[i + 1].Cells[$"phase_b_{side}"].Value = phaseB;
            i += 2;
          }
          else if (i + 2 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 2].Cells[$"breaker_{side}"].Value as string == "3") {
            phaseA = ConvertAtoVA(phaseA, 3, LINE_VOLTAGE_COMBOBOX.SelectedItem as string);
            phaseB = phaseA;
            phaseC = phaseA;
            if (phaseA == "-1" || phaseB == "-1" || phaseC == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side}, phase_b_{side}, phase_c_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value = phaseA;
            PANEL_GRID.Rows[i + 1].Cells[$"phase_b_{side}"].Value = phaseB;
            PANEL_GRID.Rows[i + 2].Cells[$"phase_c_{side}"].Value = phaseC;
            i += 3;
          }
          else {
            phaseA = ConvertAtoVA(phaseA, 1, PHASE_VOLTAGE_COMBOBOX.SelectedItem as string);
            if (phaseA == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value = phaseA;
            i++;
          }
        }
        else if (!String.IsNullOrEmpty(phaseB) && phaseB.EndsWith("A")) {
          if (i + 1 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 1].Cells[$"breaker_{side}"].Value as string == "2") {
            phaseB = ConvertAtoVA(phaseB, 2, LINE_VOLTAGE_COMBOBOX.SelectedItem as string);
            phaseC = phaseB;
            if (phaseB == "-1" || phaseC == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side} and phase_b_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value = phaseB;
            PANEL_GRID.Rows[i + 1].Cells[$"phase_c_{side}"].Value = phaseC;
            i += 2;
          }
          else if (i + 2 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 2].Cells[$"breaker_{side}"].Value as string == "3") {
            phaseB = ConvertAtoVA(phaseB, 3, LINE_VOLTAGE_COMBOBOX.SelectedItem as string);
            phaseC = phaseB;
            phaseA = phaseB;
            if (phaseB == "-1" || phaseC == "-1" || phaseA == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side}, phase_b_{side}, phase_c_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value = phaseB;
            PANEL_GRID.Rows[i + 1].Cells[$"phase_c_{side}"].Value = phaseC;
            PANEL_GRID.Rows[i + 2].Cells[$"phase_a_{side}"].Value = phaseA;
            i += 3;
          }
          else {
            phaseB = ConvertAtoVA(phaseB, 1, PHASE_VOLTAGE_COMBOBOX.SelectedItem as string);
            if (phaseB == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value = phaseB;
            i++;
          }
        }
        else if (!String.IsNullOrEmpty(phaseC) && phaseC.EndsWith("A")) {
          if (i + 1 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 1].Cells[$"breaker_{side}"].Value as string == "2") {
            phaseC = ConvertAtoVA(phaseC, 2, LINE_VOLTAGE_COMBOBOX.SelectedItem as string);
            phaseA = phaseC;
            if (phaseC == "-1" || phaseA == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side} and phase_b_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value = phaseC;
            PANEL_GRID.Rows[i + 1].Cells[$"phase_a_{side}"].Value = phaseA;
            i += 2;
          }
          else if (i + 2 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 2].Cells[$"breaker_{side}"].Value as string == "3") {
            phaseC = ConvertAtoVA(phaseC, 3, LINE_VOLTAGE_COMBOBOX.SelectedItem as string);
            phaseA = phaseC;
            phaseB = phaseC;
            if (phaseC == "-1" || phaseA == "-1" || phaseB == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side}, phase_b_{side}, phase_c_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value = phaseC;
            PANEL_GRID.Rows[i + 1].Cells[$"phase_a_{side}"].Value = phaseA;
            PANEL_GRID.Rows[i + 2].Cells[$"phase_b_{side}"].Value = phaseB;
            i += 3;
          }
          else {
            phaseC = ConvertAtoVA(phaseC, 1, PHASE_VOLTAGE_COMBOBOX.SelectedItem as string);
            if (phaseC == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value = phaseC;
            i++;
          }
        }
        else {
          i += 1;
        }
      }
    }

    private void ConvertAToVaBySide1Ph(string side) {
      int i = 0;
      while (i < PANEL_GRID.Rows.Count) {
        string phaseA = PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value?.ToString().ToUpper().Replace("\r", "");
        string phaseB = PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value?.ToString().ToUpper().Replace("\r", "");

        if (!String.IsNullOrEmpty(phaseA) && phaseA.EndsWith("A")) {
          if (i + 1 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 1].Cells[$"breaker_{side}"].Value as string == "2") {
            phaseA = ConvertAtoVA(phaseA, 2, LINE_VOLTAGE_COMBOBOX.SelectedItem as string);
            phaseB = phaseA;
            if (phaseA == "-1" || phaseB == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side} and phase_b_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value = phaseA;
            PANEL_GRID.Rows[i + 1].Cells[$"phase_b_{side}"].Value = phaseB;
            i += 2;
          }
          else {
            phaseA = ConvertAtoVA(phaseA, 1, PHASE_VOLTAGE_COMBOBOX.SelectedItem as string);
            if (phaseA == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value = phaseA;
            i++;
          }
        }
        else if (!String.IsNullOrEmpty(phaseB) && phaseB.EndsWith("A")) {
          if (i + 1 < PANEL_GRID.Rows.Count && PANEL_GRID.Rows[i + 1].Cells[$"breaker_{side}"].Value as string == "2") {
            phaseB = ConvertAtoVA(phaseB, 2, LINE_VOLTAGE_COMBOBOX.SelectedItem as string);
            phaseA = phaseB;
            if (phaseA == "-1" || phaseB == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side} and phase_b_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value = phaseB;
            PANEL_GRID.Rows[i + 1].Cells[$"phase_a_{side}"].Value = phaseA;
            i += 2;
          }
          else {
            phaseB = ConvertAtoVA(phaseB, 1, PHASE_VOLTAGE_COMBOBOX.SelectedItem as string);
            if (phaseB == "-1") {
              return;
            }
            // set values of PANEL_GRID phase_a_{side}
            PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value = phaseB;
            i++;
          }
        }
        else {
          i += 1;
        }
      }

    }

    private void A_TO_VA_BUTTON_Click(object sender, EventArgs e) {
      mainForm.SavePanelDataToLocalJsonFile();
      if (is3Ph) {
        ConvertAToVaBySide3Ph("left");
        ConvertAToVaBySide3Ph("right");
      }
      else {
        ConvertAToVaBySide1Ph("left");
        ConvertAToVaBySide1Ph("right");
      }
    }

    private void PHASE_COMBOBOX_SelectedIndexChanged(object sender, EventArgs e) {
      if (isLoading) { return; }
      bool was3Ph = is3Ph;
      is3Ph = PHASE_COMBOBOX.Text == "3";
      string side = "left";
      if (is3Ph) {
        WIRE_COMBOBOX.SelectedIndex = 1;
        AddOrRemovePanelGridColumns(is3Ph);
        ChangeSizeOfPhaseColumns(is3Ph);
        AddPhaseSumColumn(is3Ph);
        for (int j = 0; j < 2; j++) {
          for (int i = 0; i < PANEL_GRID.Rows.Count; i++) {
            if (i % 6 == 2) {
              string aToC = PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value as string;
              PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value = aToC;
              PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value = "";
            }
            else if (i % 6 == 3) {
              string bToA = PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value as string;
              PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value = bToA;
              PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value = "";
            }
            else if (i % 6 == 4) {
              string aToB = PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value as string;
              PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value = aToB;
              PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value = "";
            }
            else if (i % 6 == 5) {
              string bToC = PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value as string;
              PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value = bToC;
              PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value = "";
            }
          }
          side = "right";
        }
      } else if (PANEL_GRID.Columns.Contains($"phase_c_{side}")) {
        for (int j = 0; j < 2; j++) {
          for (int i = 0; i < PANEL_GRID.Rows.Count; i++) {
            if (i % 6 == 2) {
              string cToA = PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value as string;
              PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value = cToA;
              PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value = "";
            }
            else if (i % 6 == 3) {
              string aToB = PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value as string;
              PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value = aToB;
              PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value = "";
            }
            else if (i % 6 == 4) {
              string bToA = PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value as string;
              PANEL_GRID.Rows[i].Cells[$"phase_a_{side}"].Value = bToA;
              PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value = "";
            }
            else if (i % 6 == 5) {
              string cToB = PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value as string;
              PANEL_GRID.Rows[i].Cells[$"phase_b_{side}"].Value = cToB;
              PANEL_GRID.Rows[i].Cells[$"phase_c_{side}"].Value = "";
            }
          }
          side = "right";
        }
        WIRE_COMBOBOX.SelectedIndex = 0;
        AddOrRemovePanelGridColumns(is3Ph);
        ChangeSizeOfPhaseColumns(is3Ph);
        AddPhaseSumColumn(is3Ph);
      }
      // Toggles the distribution section checkbox.
      // This needs to happen to refresh the colors correctly as part of a separate event.
      DISTRIBUTION_SECTION_CHECKBOX.Checked = !DISTRIBUTION_SECTION_CHECKBOX.Checked;
      DISTRIBUTION_SECTION_CHECKBOX.Checked = !DISTRIBUTION_SECTION_CHECKBOX.Checked;

      SetWarnings();
    }

    private void PHASE_WARNING_LABEL_MouseHover(object sender, EventArgs e) {
      string message = "";
      if (!String.IsNullOrEmpty(FED_FROM_TEXTBOX.Text)) {
        message += "Phase and wire cannot be altered when panel is fed from another panel.\n";
      }
      if (this.contains3PhEquip) {
        message += "Phase and wire cannot be altered when panel has 3-pole breakers.";
      }
      if (!String.IsNullOrEmpty(message)) {
        System.Windows.Forms.ToolTip toolTip = new System.Windows.Forms.ToolTip();
        toolTip.SetToolTip(PHASE_WARNING_LABEL, message);
      }
    }

    private void HIGH_LEG_WARNING_LEFT_LABEL_MouseHover(object sender, EventArgs e) {
      string message = "High-leg phase cannot have single-phase, single-pole breaker.";
      System.Windows.Forms.ToolTip toolTip = new System.Windows.Forms.ToolTip();
      toolTip.SetToolTip(HIGH_LEG_WARNING_LEFT_LABEL, message);
    }
    
    private void HIGH_LEG_WARNING_RIGHT_LABEL_MouseHover(object sender, EventArgs e) {
      string message = "High-leg phase cannot have single-phase, single-pole breaker.";
      System.Windows.Forms.ToolTip toolTip = new System.Windows.Forms.ToolTip();
      toolTip.SetToolTip(HIGH_LEG_WARNING_RIGHT_LABEL, message);
    }
  }

  public class PanelItem {
    public string Description { get; set; }
    public double Wattage { get; set; }
    public int Poles { get; set; }
  }
}