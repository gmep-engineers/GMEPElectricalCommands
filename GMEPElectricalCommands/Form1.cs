﻿using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using static OfficeOpenXml.ExcelErrorValue;
using Autodesk.AutoCAD.GraphicsInterface;

namespace AutoCADCommands
{
  public partial class Form1 : Form
  {
    private MyCommands myCommandsInstance;

    public Form1(MyCommands myCommands)
    {
      myCommandsInstance = myCommands;

      InitializeComponent();
      Load_Panel_Data();
      Listen_For_New_Rows();
      Remove_Column_Header_Sorting();

      PANEL_GRID.Rows.AddCopies(0, 21);
      PANEL_GRID.AllowUserToAddRows = false;
      PANEL_GRID.KeyDown += new KeyEventHandler(this.PANEL_GRID_KeyDown);
      PANEL_GRID.CellBeginEdit += new DataGridViewCellCancelEventHandler(this.PANEL_GRID_CellBeginEdit);
      PANEL_GRID.CellValueChanged += new DataGridViewCellEventHandler(this.PANEL_GRID_CellValueChanged);
      PHASE_SUM_GRID.CellValueChanged += new DataGridViewCellEventHandler(this.PHASE_SUM_GRID_CellValueChanged);
      PANEL_GRID.CellFormatting += PANEL_GRID_CellFormatting;

      Add_Rows_To_DataGrids();
      Set_Default_Form_Values();
      Deselect_Cells();
    }

    private void Add_Rows_To_DataGrids()
    {
      // Datagrids
      PHASE_SUM_GRID.Rows.Add("0", "0");
      TOTAL_VA_GRID.Rows.Add("0");
      LCL_GRID.Rows.Add("0", "0");
      TOTAL_OTHER_LOAD_GRID.Rows.Add("0");
      PANEL_LOAD_GRID.Rows.Add("0");
      FEEDER_AMP_GRID.Rows.Add("0");
    }

    private void Load_Panel_Data()
    {
      List<Dictionary<string, object>> saveData = Retrieve_Saved_Data();

      foreach (Dictionary<string, object> panel in saveData)
      {
        string panelName = panel["panel"].ToString();
        LOAD_PANEL_COMBOBOX.Items.Add(panelName);
      }
    }

    private List<Dictionary<string, object>> Retrieve_Saved_Data()
    {
      List<Dictionary<string, object>> saveData = new List<Dictionary<string, object>>();

      Autodesk.AutoCAD.ApplicationServices.Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Autodesk.AutoCAD.DatabaseServices.Database acCurDb = acDoc.Database;
      string jsonDataKey = "JsonSaveData";

      using (Autodesk.AutoCAD.DatabaseServices.Transaction tr = acCurDb.TransactionManager.StartTransaction())
      {
        Autodesk.AutoCAD.DatabaseServices.DBDictionary nod = (Autodesk.AutoCAD.DatabaseServices.DBDictionary)tr.GetObject(acCurDb.NamedObjectsDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

        if (nod.Contains(jsonDataKey))
        {
          Autodesk.AutoCAD.DatabaseServices.DBDictionary userDict = (Autodesk.AutoCAD.DatabaseServices.DBDictionary)tr.GetObject(nod.GetAt(jsonDataKey), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
          Autodesk.AutoCAD.DatabaseServices.Xrecord xRecord = (Autodesk.AutoCAD.DatabaseServices.Xrecord)tr.GetObject(userDict.GetAt("SaveData"), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
          Autodesk.AutoCAD.DatabaseServices.ResultBuffer rb = xRecord.Data;
          if (rb != null)
          {
            foreach (Autodesk.AutoCAD.DatabaseServices.TypedValue tv in rb)
            {
              if (tv.TypeCode == (int)Autodesk.AutoCAD.DatabaseServices.DxfCode.Text)
              {
                saveData = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(tv.Value.ToString());
              }
            }
          }
        }
      }

      return saveData;
    }

    private void Store_Data(List<Dictionary<string, object>> saveData)
    {
      Autodesk.AutoCAD.ApplicationServices.Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
      Autodesk.AutoCAD.DatabaseServices.Database acCurDb = acDoc.Database;
      string jsonDataKey = "JsonSaveData";

      using (Autodesk.AutoCAD.DatabaseServices.Transaction tr = acCurDb.TransactionManager.StartTransaction())
      {
        Autodesk.AutoCAD.DatabaseServices.DBDictionary nod = (Autodesk.AutoCAD.DatabaseServices.DBDictionary)tr.GetObject(acCurDb.NamedObjectsDictionaryId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);

        if (!nod.Contains(jsonDataKey))
        {
          Autodesk.AutoCAD.DatabaseServices.DBDictionary userDict = new Autodesk.AutoCAD.DatabaseServices.DBDictionary();
          nod.UpgradeOpen();
          nod.SetAt(jsonDataKey, userDict);
          tr.AddNewlyCreatedDBObject(userDict, true);

          Autodesk.AutoCAD.DatabaseServices.Xrecord xRecord = new Autodesk.AutoCAD.DatabaseServices.Xrecord();
          Autodesk.AutoCAD.DatabaseServices.ResultBuffer rb = new Autodesk.AutoCAD.DatabaseServices.ResultBuffer(new Autodesk.AutoCAD.DatabaseServices.TypedValue((int)Autodesk.AutoCAD.DatabaseServices.DxfCode.Text, JsonConvert.SerializeObject(saveData, Formatting.Indented)));
          xRecord.Data = rb;
          userDict.SetAt("SaveData", xRecord);
          tr.AddNewlyCreatedDBObject(xRecord, true);
        }

        tr.Commit();
      }
    }

    private object oldValue;

    private void PANEL_GRID_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
    {
      oldValue = PANEL_GRID.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
    }

    private void PANEL_GRID_KeyDown(object sender, KeyEventArgs e)
    {
      if (e.Control && e.KeyCode == Keys.V)
      {
        // Get text from clipboard
        string text = Clipboard.GetText();

        if (!string.IsNullOrEmpty(text))
        {
          // Split clipboard text into lines
          string[] lines = text.Split('\n');

          if (lines.Length > 0 && string.IsNullOrWhiteSpace(lines[lines.Length - 1]))
          {
            Array.Resize(ref lines, lines.Length - 1);
          }

          // Get start cell for pasting
          int rowIndex = PANEL_GRID.CurrentCell.RowIndex;
          int colIndex = PANEL_GRID.CurrentCell.ColumnIndex;
          int startRowIndex = PANEL_GRID.CurrentCell.RowIndex;

          // Paste each line into a row
          foreach (string line in lines)
          {
            string[] parts = line.Split('\t');

            for (int i = 0; i < parts.Length; i++)
            {
              if (rowIndex < PANEL_GRID.RowCount && colIndex + i < PANEL_GRID.ColumnCount)
              {
                try
                {
                  PANEL_GRID[colIndex + i, rowIndex].Value = parts[i];
                }
                catch (FormatException)
                {
                  // Handle format exception

                  // Set to default value
                  PANEL_GRID[colIndex + i, rowIndex].Value = 0;

                  // Or leave cell blank
                  //dataGridView1[colIndex + i, rowIndex].Value = DBNull.Value;

                  // Or notify user
                  MessageBox.Show("Invalid format in cell!");
                }
              }
            }

            rowIndex++;
          }
          // Reset row index after loop
          rowIndex = startRowIndex;
        }

        e.Handled = true;
      }
      // Check if Ctrl+C was pressed
      else if (e.Control && e.KeyCode == Keys.C)
      {
        StringBuilder copiedText = new StringBuilder();

        // Loop through selected cells
        foreach (DataGridViewCell cell in PANEL_GRID.SelectedCells)
        {
          copiedText.AppendLine(cell.Value?.ToString() ?? string.Empty);
        }

        // Loop through selected rows
        foreach (DataGridViewRow row in PANEL_GRID.SelectedRows)
        {
          List<string> cellValues = new List<string>();
          foreach (DataGridViewCell cell in row.Cells)
          {
            if (cell.Selected)
            {
              cellValues.Add(cell.Value?.ToString() ?? string.Empty);
            }
          }
          if (cellValues.Count > 0)
          {
            copiedText.AppendLine(string.Join("\t", cellValues));
          }
        }

        if (copiedText.Length > 0)
        {
          Clipboard.SetText(copiedText.ToString());
        }

        e.Handled = true;
      }

      // Existing code for handling the Delete key
      else if (e.KeyCode == Keys.Delete)
      {
        foreach (DataGridViewCell cell in PANEL_GRID.SelectedCells)
        {
          cell.Value = null;
        }
        e.Handled = true;
      }
    }

    private void PHASE_SUM_GRID_CellValueChanged(object sender, DataGridViewCellEventArgs e)
    {
      // Check if the modified cell is in row 0 and column 0 or 1
      if (e.RowIndex == 0 && (e.ColumnIndex == 0 || e.ColumnIndex == 1))
      {
        // Retrieve the values from PHASE_SUM_GRID, row 0, column 0 and 1
        double value1 = Convert.ToDouble(PHASE_SUM_GRID.Rows[0].Cells[0].Value);
        double value2 = Convert.ToDouble(PHASE_SUM_GRID.Rows[0].Cells[1].Value);

        // Perform the sum (checking for null values)
        double sum = value1 + value2;

        if (LARGEST_LCL_CHECKBOX.Checked && TOTAL_OTHER_LOAD_GRID[0, 0].Value != null)
        {
          calculate_totalva_panelload_feederamps_lcl(sum);
        }
        else
        {
          calculate_totalva_panelload_feederamps_regular(value1, value2, sum);
        }
      }
    }

    private void calculate_totalva_panelload_feederamps_lcl(double sum)
    {
      // Update total VA
      TOTAL_VA_GRID.Rows[0].Cells[0].Value = Math.Round(sum, 0); // Rounded to 0 decimal places

      // Update panel load grid including the other load value
      double otherLoad = Convert.ToDouble(TOTAL_OTHER_LOAD_GRID[0, 0].Value);
      PANEL_LOAD_GRID.Rows[0].Cells[0].Value = Math.Round((otherLoad + sum) / 1000, 1); // Rounded to 1 decimal place

      // Update feeder amps
      double panelLoadValue = Convert.ToDouble(Convert.ToDouble(TOTAL_VA_GRID.Rows[0].Cells[0].Value) + otherLoad) / (120 * PHASE_SUM_GRID.ColumnCount);
      FEEDER_AMP_GRID.Rows[0].Cells[0].Value = Math.Round(panelLoadValue, 1); // Rounded to 1 decimal place
    }

    private void calculate_totalva_panelload_feederamps_regular(double value1, double value2, double sum)
    {
      // Update total VA and panel load without other loads
      TOTAL_VA_GRID.Rows[0].Cells[0].Value = Math.Round(sum, 0); // Rounded to 1 decimal place
      PANEL_LOAD_GRID.Rows[0].Cells[0].Value = Math.Round(sum / 1000, 1); // Rounded to 1 decimal place

      // Update feeder amps
      double maxVal = Math.Max(Convert.ToDouble(value1), Convert.ToDouble(value2));

      object lineVoltageObj = LINE_VOLTAGE_COMBOBOX.SelectedItem;
      if (lineVoltageObj != null)
      {
        double lineVoltage = Convert.ToDouble(lineVoltageObj);
        if (lineVoltage != 0)
        {
          double panelLoadValue = maxVal / lineVoltage;
          FEEDER_AMP_GRID.Rows[0].Cells[0].Value = Math.Round(panelLoadValue, 1); // Rounded to 1 decimal place
        }
      }
    }

    private void PANEL_GRID_CellValueChanged(object sender, DataGridViewCellEventArgs e)
    {
      // Function to parse and sum a cell's value
      double ParseAndSumCell(string cellValue)
      {
        double sum = 0;
        if (!string.IsNullOrEmpty(cellValue))
        {
          var parts = cellValue.Split(',');
          foreach (var part in parts)
          {
            if (double.TryParse(part, out double value))
            {
              sum += value;
            }
          }
        }
        return sum;
      }

      // Check if the modified cell is in column 2 or 8
      if (e.ColumnIndex == 1 || e.ColumnIndex == 7)
      {
        double sum = 0;
        foreach (DataGridViewRow row in PANEL_GRID.Rows)
        {
          if (row.Cells[1].Value != null)
            sum += ParseAndSumCell(row.Cells[1].Value.ToString());

          if (row.Cells[7].Value != null)
            sum += ParseAndSumCell(row.Cells[7].Value.ToString());
        }

        // Update the sum in dataGridView2, row 0, column 0
        PHASE_SUM_GRID.Rows[0].Cells[0].Value = sum;
      }

      // Check if the modified cell is in column 3 or 9
      if (e.ColumnIndex == 2 || e.ColumnIndex == 8)
      {
        double sum = 0;
        foreach (DataGridViewRow row in PANEL_GRID.Rows)
        {
          if (row.Cells[2].Value != null)
            sum += ParseAndSumCell(row.Cells[2].Value.ToString());

          if (row.Cells[8].Value != null)
            sum += ParseAndSumCell(row.Cells[8].Value.ToString());
        }

        // Update the sum in dataGridView2, row 0, column 1
        PHASE_SUM_GRID.Rows[0].Cells[1].Value = sum;
      }

      object newValue = PANEL_GRID.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
      // oldValue = null;  // Not sure if you use this, so leaving it commented
    }

    private void PANEL_GRID_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
    {
      // Check if the current cell is the one being formatted and the DataGridView doesn't have focus
      if (this.PANEL_GRID.CurrentCell != null
          && e.RowIndex == this.PANEL_GRID.CurrentCell.RowIndex
          && e.ColumnIndex == this.PANEL_GRID.CurrentCell.ColumnIndex
          && !this.PANEL_GRID.Focused)
      {
        // Change the back color and fore color to make the current cell less noticeable
        e.CellStyle.SelectionBackColor = e.CellStyle.BackColor;
        e.CellStyle.SelectionForeColor = e.CellStyle.ForeColor;
      }
    }

    private void Deselect_Cells()
    {
      PHASE_SUM_GRID.DefaultCellStyle.SelectionBackColor = PHASE_SUM_GRID.DefaultCellStyle.BackColor;
      PHASE_SUM_GRID.DefaultCellStyle.SelectionForeColor = PHASE_SUM_GRID.DefaultCellStyle.ForeColor;
      TOTAL_VA_GRID.DefaultCellStyle.SelectionBackColor = TOTAL_VA_GRID.DefaultCellStyle.BackColor;
      TOTAL_VA_GRID.DefaultCellStyle.SelectionForeColor = TOTAL_VA_GRID.DefaultCellStyle.ForeColor;
      LCL_GRID.DefaultCellStyle.SelectionBackColor = LCL_GRID.DefaultCellStyle.BackColor;
      LCL_GRID.DefaultCellStyle.SelectionForeColor = LCL_GRID.DefaultCellStyle.ForeColor;
      TOTAL_OTHER_LOAD_GRID.DefaultCellStyle.SelectionBackColor = TOTAL_OTHER_LOAD_GRID.DefaultCellStyle.BackColor;
      TOTAL_OTHER_LOAD_GRID.DefaultCellStyle.SelectionForeColor = TOTAL_OTHER_LOAD_GRID.DefaultCellStyle.ForeColor;
      PANEL_LOAD_GRID.DefaultCellStyle.SelectionBackColor = PANEL_LOAD_GRID.DefaultCellStyle.BackColor;
      PANEL_LOAD_GRID.DefaultCellStyle.SelectionForeColor = PANEL_LOAD_GRID.DefaultCellStyle.ForeColor;
      FEEDER_AMP_GRID.DefaultCellStyle.SelectionBackColor = FEEDER_AMP_GRID.DefaultCellStyle.BackColor;
      FEEDER_AMP_GRID.DefaultCellStyle.SelectionForeColor = FEEDER_AMP_GRID.DefaultCellStyle.ForeColor;
      PANEL_GRID.ClearSelection();
    }

    private void Set_Default_Form_Values()
    {
      // Textboxes
      PANEL_NAME_INPUT.Text = "A";
      PANEL_LOCATION_INPUT.Text = "ELECTRIC ROOM";
      MAIN_INPUT.Text = "M.L.O";
      BUS_RATING_INPUT.Text = "100";

      // Comboboxes
      STATUS_COMBOBOX.SelectedIndex = 0;
      MOUNTING_COMBOBOX.SelectedIndex = 0;
      WIRE_COMBOBOX.SelectedIndex = 0;
      PHASE_COMBOBOX.SelectedIndex = 0;
      PHASE_VOLTAGE_COMBOBOX.SelectedIndex = 0;
      LINE_VOLTAGE_COMBOBOX.SelectedIndex = 0;

      // Datagrids
      PHASE_SUM_GRID.Rows[0].Cells[0].Value = "0";
      PHASE_SUM_GRID.Rows[0].Cells[1].Value = "0";
      TOTAL_VA_GRID.Rows[0].Cells[0].Value = "0";
      LCL_GRID.Rows[0].Cells[0].Value = "0";
      LCL_GRID.Rows[0].Cells[1].Value = "0";
      TOTAL_OTHER_LOAD_GRID.Rows[0].Cells[0].Value = "0";
      PANEL_LOAD_GRID.Rows[0].Cells[0].Value = "0";
      FEEDER_AMP_GRID.Rows[0].Cells[0].Value = "0";
    }

    private void Remove_Column_Header_Sorting()
    {
      foreach (DataGridViewColumn column in PANEL_GRID.Columns)
      {
        column.SortMode = DataGridViewColumnSortMode.NotSortable;
      }
    }

    private void Listen_For_New_Rows()
    {
      PANEL_GRID.RowsAdded += new DataGridViewRowsAddedEventHandler(PANEL_GRID_RowsAdded);
    }

    private void PANEL_GRID_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
    {
      for (int i = 0; i < e.RowCount; i++)
      {
        int rowIndex = e.RowIndex + i;

        // Set common column values
        PANEL_GRID.Rows[rowIndex].Cells[3].Value = "20";
        PANEL_GRID.Rows[rowIndex].Cells[4].Value = ((rowIndex + 1) * 2) - 1;
        PANEL_GRID.Rows[rowIndex].Cells[5].Value = (rowIndex + 1) * 2;
        PANEL_GRID.Rows[rowIndex].Cells[6].Value = "20";

        // Zig-zag pattern for columns 2, 3, 8, and 9
        if ((rowIndex + 1) % 2 == 1) // Odd rows
        {
          PANEL_GRID.Rows[rowIndex].Cells[1].Style.BackColor = Color.LightGray; // Column 2
          PANEL_GRID.Rows[rowIndex].Cells[7].Style.BackColor = Color.LightGray; // Column 8
        }
        else // Even rows
        {
          PANEL_GRID.Rows[rowIndex].Cells[2].Style.BackColor = Color.LightGray; // Column 3
          PANEL_GRID.Rows[rowIndex].Cells[8].Style.BackColor = Color.LightGray; // Column 9
        }
      }
    }

    private void ADD_ROW_BUTTON_CLICK(object sender, EventArgs e)
    {
      PANEL_GRID.Rows.Add();
    }

    private void DELETE_ROW_BUTTON_CLICK(object sender, EventArgs e)
    {
      if (PANEL_GRID.Rows.Count > 0)
      {
        PANEL_GRID.Rows.RemoveAt(PANEL_GRID.Rows.Count - 1);
      }
    }

    private void CREATE_PANEL_BUTTON_CLICK(object sender, EventArgs e)
    {
      Dictionary<string, object> panelDataList = Retrieve_Data_From_Modal();
      myCommandsInstance.Create_Panel(panelDataList);
      this.Close();
    }

    private void Print_Panels(List<Dictionary<string, object>> panels)
    {
      foreach (Dictionary<string, object> panel in panels)
      {
        string jsonFormattedString = JsonConvert.SerializeObject(panel, Formatting.Indented);
        Console.WriteLine(jsonFormattedString);
      }
    }

    private Dictionary<string, object> Retrieve_Data_From_Modal()
    {
      // Create a new panel
      Dictionary<string, object> panel = new Dictionary<string, object>();

      // Get the value from the main input
      string mainInput = MAIN_INPUT.Text;

      // Check if the value contains the word "amp" or "AMP"
      if (mainInput.ToLower().Contains("amp"))
      {
        mainInput = mainInput.ToUpper().Replace("A ", "AMP ").Replace(" A", " AMP");
      }
      // Check if the value is just a number
      else if (IsDigitsOnly(mainInput.Replace(" ", "")))
      {
        mainInput = mainInput + " AMP";
      }
      else
      {
        string[] parts = mainInput.Split(' ');
        if (parts.Length > 1 && IsDigitsOnly(parts[0]) && parts[1].ToLower() == "a")
        {
          mainInput = parts[0] + " AMP";
        }
        // Add any other conditions here if needed
      }

      // Add the processed main input to the panel dictionary
      panel.Add("main", mainInput.ToUpper());

      string GetComboBoxValue(ComboBox comboBox)
      {
        if (comboBox.SelectedItem != null)
        {
          return comboBox.SelectedItem.ToString().ToUpper();
        }
        else if (!string.IsNullOrEmpty(comboBox.Text))
        {
          return comboBox.Text.ToUpper();
        }
        else
        {
          return ""; // Default value or you can return null
        }
      }

      // Add simple values in uppercase
      panel.Add("panel", "'" + PANEL_NAME_INPUT.Text.ToUpper() + "'");
      panel.Add("location", PANEL_LOCATION_INPUT.Text.ToUpper());
      panel.Add("voltage1", GetComboBoxValue(LINE_VOLTAGE_COMBOBOX));
      panel.Add("voltage2", GetComboBoxValue(PHASE_VOLTAGE_COMBOBOX));
      panel.Add("phase", GetComboBoxValue(PHASE_COMBOBOX));
      panel.Add("wire", GetComboBoxValue(WIRE_COMBOBOX));
      panel.Add("mounting", GetComboBoxValue(MOUNTING_COMBOBOX));
      panel.Add("existing", GetComboBoxValue(STATUS_COMBOBOX));

      // Add datagrid values in uppercase
      // Assuming that these grids are DataGridViews and the specific cells mentioned are not null or empty
      panel.Add("subtotal_a", PHASE_SUM_GRID.Rows[0].Cells[0].Value.ToString().ToUpper());
      panel.Add("subtotal_b", PHASE_SUM_GRID.Rows[0].Cells[1].Value.ToString().ToUpper());
      panel.Add("subtotal_c", "0"); // Set to "0" as per the requirement, no need to convert
      panel.Add("total_va", TOTAL_VA_GRID.Rows[0].Cells[0].Value.ToString().ToUpper());
      panel.Add("lcl", LCL_GRID.Rows[0].Cells[1].Value.ToString().ToUpper());
      panel.Add("lcl_125", LCL_GRID.Rows[0].Cells[0].Value.ToString().ToUpper());
      panel.Add("total_other_load", TOTAL_OTHER_LOAD_GRID.Rows[0].Cells[0].Value.ToString().ToUpper());
      panel.Add("kva", PANEL_LOAD_GRID.Rows[0].Cells[0].Value.ToString().ToUpper());
      panel.Add("feeder_amps", FEEDER_AMP_GRID.Rows[0].Cells[0].Value.ToString().ToUpper());

      // Add "A" to the bus rating value if it consists of digits only, then convert to uppercase
      string busRatingInput = BUS_RATING_INPUT.Text;
      if (IsDigitsOnly(busRatingInput))
      {
        busRatingInput += "A"; // append "A" if the input is numeric
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
      List<string> breaker_left = new List<string>();
      List<string> breaker_right = new List<string>();
      List<string> circuit_left = new List<string>();
      List<string> circuit_right = new List<string>();

      for (int i = 0; i < PANEL_GRID.Rows.Count; i++)
      {
        string descriptionLeftValue = PANEL_GRID.Rows[i].Cells["description_left"].Value?.ToString().ToUpper() ?? "SPACE";
        string breakerLeftValue = PANEL_GRID.Rows[i].Cells["breaker_left"].Value?.ToString().ToUpper() ?? "";
        string descriptionRightValue = PANEL_GRID.Rows[i].Cells["description_right"].Value?.ToString().ToUpper() ?? "SPACE";
        string breakerRightValue = PANEL_GRID.Rows[i].Cells["breaker_right"].Value?.ToString().ToUpper() ?? "";
        string circuitRightValue = PANEL_GRID.Rows[i].Cells["circuit_right"].Value?.ToString().ToUpper() ?? "";
        string circuitLeftValue = PANEL_GRID.Rows[i].Cells["circuit_left"].Value?.ToString().ToUpper() ?? "";
        string phaseALeftValue = PANEL_GRID.Rows[i].Cells["phase_a_left"].Value?.ToString() ?? "0";
        string phaseBLeftValue = PANEL_GRID.Rows[i].Cells["phase_b_left"].Value?.ToString() ?? "0";
        string phaseARightValue = PANEL_GRID.Rows[i].Cells["phase_a_right"].Value?.ToString() ?? "0";
        string phaseBRightValue = PANEL_GRID.Rows[i].Cells["phase_b_right"].Value?.ToString() ?? "0";

        // Checks for Left Side
        bool hasCommaInPhaseLeft = phaseALeftValue.Contains(",") || phaseBLeftValue.Contains(",");
        bool shouldDuplicateLeft = hasCommaInPhaseLeft;

        // Checks for Right Side
        bool hasCommaInPhaseRight = phaseARightValue.Contains(",") || phaseBRightValue.Contains(",");
        bool shouldDuplicateRight = hasCommaInPhaseRight;

        // Handling Phase A Left
        if (phaseALeftValue.Contains(","))
        {
          var splitValues = phaseALeftValue.Split(',').Select(str => str.Trim()).ToArray();
          phase_a_left.AddRange(splitValues);
        }
        else
        {
          phase_a_left.Add(phaseALeftValue);
          phase_a_left.Add("0"); // Default value
        }

        // Handling Phase B Left
        if (phaseBLeftValue.Contains(","))
        {
          var splitValues = phaseBLeftValue.Split(',').Select(str => str.Trim()).ToArray();
          phase_b_left.AddRange(splitValues);
        }
        else
        {
          phase_b_left.Add(phaseBLeftValue);
          phase_b_left.Add("0"); // Default value
        }

        // Handling Phase A Right
        if (phaseARightValue.Contains(","))
        {
          var splitValues = phaseARightValue.Split(',').Select(str => str.Trim()).ToArray();
          phase_a_right.AddRange(splitValues);
        }
        else
        {
          phase_a_right.Add(phaseARightValue);
          phase_a_right.Add("0"); // Default value
        }

        // Handling Phase B Right
        if (phaseBRightValue.Contains(","))
        {
          var splitValues = phaseBRightValue.Split(',').Select(str => str.Trim()).ToArray();
          phase_b_right.AddRange(splitValues);
        }
        else
        {
          phase_b_right.Add(phaseBRightValue);
          phase_b_right.Add("0"); // Default value
        }

        if (descriptionLeftValue.Contains(","))
        {
          // If it contains a comma, split and add both values
          var splitValues = descriptionLeftValue.Split(',')
                                                .Select(str => str.Trim())
                                                .ToArray();
          description_left.AddRange(splitValues);
          circuit_left.Add(circuitLeftValue + "A");
          circuit_left.Add(circuitLeftValue + "B");
        }
        else
        {
          description_left.Add(descriptionLeftValue);
          description_left.Add(shouldDuplicateLeft ? descriptionLeftValue : "SPACE");

          if (shouldDuplicateLeft)
          {
            circuit_left.Add(circuitLeftValue + "A");
            circuit_left.Add(circuitLeftValue + "B");
          }
          else
          {
            circuit_left.Add(circuitLeftValue);
            circuit_left.Add("");
          }
        }

        if (breakerLeftValue.Contains(","))
        {
          // If it contains a comma, split and add both values
          var splitValues = breakerLeftValue.Split(',')
                                                .Select(str => str.Trim())
                                                .ToArray();
          breaker_left.AddRange(splitValues);
        }
        else
        {
          breaker_left.Add(breakerLeftValue);
          breaker_left.Add(shouldDuplicateLeft ? breakerLeftValue : "");
        }

        if (descriptionRightValue.Contains(","))
        {
          // If it contains a comma, split and add both values
          var splitValues = descriptionRightValue.Split(',')
                                                .Select(str => str.Trim())
                                                .ToArray();
          description_right.AddRange(splitValues);
          circuit_right.Add(circuitRightValue + "A");
          circuit_right.Add(circuitRightValue + "B");
        }
        else
        {
          description_right.Add(descriptionRightValue);
          description_right.Add(shouldDuplicateRight ? descriptionRightValue : "SPACE");

          if (shouldDuplicateRight)
          {
            circuit_right.Add(circuitRightValue + "A");
            circuit_right.Add(circuitRightValue + "B");
          }
          else
          {
            circuit_right.Add(circuitRightValue);
            circuit_right.Add("");
          }
        }

        if (breakerRightValue.Contains(","))
        {
          // If it contains a comma, split and add both values
          var splitValues = breakerRightValue.Split(',')
                                                .Select(str => str.Trim())
                                                .ToArray();
          breaker_right.AddRange(splitValues);
        }
        else
        {
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
      panel.Add("breaker_left", breaker_left);
      panel.Add("breaker_right", breaker_right);
      panel.Add("circuit_left", circuit_left);
      panel.Add("circuit_right", circuit_right);

      return panel;
    }

    private bool IsDigitsOnly(string str)
    {
      foreach (char c in str)
      {
        if (c < '0' || c > '9')
          return false;
      }
      return true;
    }

    private void LARGEST_LCL_CHECKBOX_CheckedChanged(object sender, EventArgs e)
    {
      calculate_lcl_otherload_panelload_feederamps();
    }

    private void LARGEST_LCL_INPUT_TextChanged(object sender, EventArgs e)
    {
      calculate_lcl_otherload_panelload_feederamps();
    }

    private void calculate_lcl_otherload_panelload_feederamps()
    {
      if (LARGEST_LCL_CHECKBOX.Checked)
      {
        // 1. Check if the textbox is empty
        string largestLclInputText = LARGEST_LCL_INPUT.Text;
        if (string.IsNullOrEmpty(largestLclInputText))
        {
          return;
        }

        // 2. If it has a number, put that number in col 0 row 0 of "LCL_GRID".
        double largestLclInputValue;
        if (double.TryParse(largestLclInputText, out largestLclInputValue))
        {
          LCL_GRID[0, 0].Value = largestLclInputValue;

          // 3. Multiply by 125% and put that in col 1 row 0 of "LCL_GRID".
          double value125 = largestLclInputValue * 1.25;
          LCL_GRID[1, 0].Value = value125;

          // 4. Subtract value in col 1 row 0 from value in col 0 row 0 of "LCL_GRID".
          double difference = value125 - largestLclInputValue;

          // 5. Put that number in col 0 row 0 of "TOTAL_OTHER_LOAD_GRID".
          TOTAL_OTHER_LOAD_GRID[0, 0].Value = difference;

          // 6. Add to the value in "PANEL_LOAD_GRID".
          double panelLoad = Convert.ToDouble(TOTAL_VA_GRID[0, 0].Value);
          double total_KVA = (panelLoad + difference) / 1000;
          PANEL_LOAD_GRID[0, 0].Value = Math.Round(total_KVA, 1);

          // 7. Divide by value in "PHASE_VOLTAGE_COMBOBOX".
          double lineVoltage = Convert.ToDouble(LINE_VOLTAGE_COMBOBOX.SelectedItem);
          if (lineVoltage != 0)
          {
            double result = (panelLoad + difference) / (lineVoltage * 2);
            FEEDER_AMP_GRID[0, 0].Value = Math.Round(result, 1);
          }
        }
        else
        {
          MessageBox.Show("Invalid number format in LARGEST_LCL_INPUT.");
        }
      }
      else
      {
        LCL_GRID[0, 0].Value = "0";
        LCL_GRID[1, 0].Value = "0";
        TOTAL_OTHER_LOAD_GRID[0, 0].Value = "0";

        // Retrieve the values from PHASE_SUM_GRID, row 0, column 0 and 1
        double value1 = Convert.ToDouble(PHASE_SUM_GRID.Rows[0].Cells[0].Value);
        double value2 = Convert.ToDouble(PHASE_SUM_GRID.Rows[0].Cells[1].Value);

        // Perform the sum (checking for null values)
        double sum = value1 + value2;

        calculate_totalva_panelload_feederamps_regular(value1, value2, sum);
      }
    }

    private void SAVE_PANEL_BUTTON_Click(object sender, EventArgs e)
    {
      save_panel_data();
    }

    private void save_panel_data()
    {
      string formattedPanelName = $"'{PANEL_NAME_INPUT.Text.ToUpper()}'";

      if (string.IsNullOrEmpty(PANEL_NAME_INPUT.Text))
      {
        MessageBox.Show("Please enter a value before proceeding.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        return;
      }

      List<Dictionary<string, object>> saveData = Retrieve_Saved_Data();
      Dictionary<string, object> currentPanelData = Retrieve_Data_From_Modal();

      bool panelExists = saveData.Any(dict => dict["panel"].ToString() == formattedPanelName);

      if (panelExists)
      {
        DialogResult result = MessageBox.Show("The panel name already exists. Do you want to overwrite the existing panel?", "Confirm", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

        switch (result)
        {
          case DialogResult.Yes:
            // Find and remove the existing panel data
            Dictionary<string, object> existingPanel = saveData.First(dict => dict["panel"].ToString() == formattedPanelName);
            saveData.Remove(existingPanel);
            saveData.Add(currentPanelData);
            break;

          case DialogResult.No:
            // Do nothing and return
            return;

          case DialogResult.Cancel:
            // Do nothing and return
            return;
        }
      }
      else
      {
        saveData.Add(currentPanelData);
        LOAD_PANEL_COMBOBOX.Items.Add(formattedPanelName); // Adding the formatted panel name to the ComboBox
      }

      Store_Data(saveData);
    }

    private void LOAD_PANEL_BUTTON_click(object sender, EventArgs e)
    {
      List<Dictionary<string, object>> saveData = Retrieve_Saved_Data();

      if (LOAD_PANEL_COMBOBOX.SelectedItem != null)
      {
        string selectedPanelName = LOAD_PANEL_COMBOBOX.SelectedItem.ToString();
        Dictionary<string, object> selectedPanelData = saveData.First(dict => dict["panel"].ToString() == selectedPanelName);

        // prompt the user if they would like to save the current panel
        DialogResult result = MessageBox.Show("Would you like to save the current panel?", "Confirm", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        if (result == DialogResult.Yes)
        {
          save_panel_data();
        }
        else if (result == DialogResult.Cancel)
        {
          return;
        }

        // Set the values in the modal
        Set_Modal_Values(selectedPanelData);
      }
    }

    private void Set_Modal_Values(Dictionary<string, object> selectedPanelData)
    {
      clear_current_modal_data();

      populate_modal_with_panel_data(selectedPanelData);
    }

    private void populate_modal_with_panel_data(Dictionary<string, object> selectedPanelData)
    {
      // Set TextBoxes
      MAIN_INPUT.Text = selectedPanelData["main"].ToString();
      PANEL_NAME_INPUT.Text = selectedPanelData["panel"].ToString().Replace("'", "");
      PANEL_LOCATION_INPUT.Text = selectedPanelData["location"].ToString();
      BUS_RATING_INPUT.Text = selectedPanelData["bus_rating"].ToString().Replace("A", "");

      // Set ComboBoxes
      STATUS_COMBOBOX.SelectedItem = selectedPanelData["existing"].ToString();
      MOUNTING_COMBOBOX.SelectedItem = selectedPanelData["mounting"].ToString();
      WIRE_COMBOBOX.SelectedItem = selectedPanelData["wire"].ToString();
      PHASE_COMBOBOX.SelectedItem = selectedPanelData["phase"].ToString();
      PHASE_VOLTAGE_COMBOBOX.SelectedItem = selectedPanelData["voltage2"].ToString();
      LINE_VOLTAGE_COMBOBOX.SelectedItem = selectedPanelData["voltage1"].ToString();

      // Set DataGridViews
      PHASE_SUM_GRID.Rows[0].Cells[0].Value = selectedPanelData["subtotal_a"].ToString();
      PHASE_SUM_GRID.Rows[0].Cells[1].Value = selectedPanelData["subtotal_b"].ToString();
      TOTAL_VA_GRID.Rows[0].Cells[0].Value = selectedPanelData["total_va"].ToString();
      LCL_GRID.Rows[0].Cells[0].Value = selectedPanelData["lcl_125"].ToString();
      LCL_GRID.Rows[0].Cells[1].Value = selectedPanelData["lcl"].ToString();
      TOTAL_OTHER_LOAD_GRID.Rows[0].Cells[0].Value = selectedPanelData["total_other_load"].ToString();
      PANEL_LOAD_GRID.Rows[0].Cells[0].Value = selectedPanelData["kva"].ToString();
      FEEDER_AMP_GRID.Rows[0].Cells[0].Value = selectedPanelData["feeder_amps"].ToString();

      List<string> multi_row_datagrid_keys = new List<string> { "description_left", "description_right", "phase_a_left", "phase_b_left", "phase_a_right", "phase_b_right", "breaker_left", "breaker_right", "circuit_left", "circuit_right" };

      int length = ((Newtonsoft.Json.Linq.JArray)selectedPanelData["description_left"]).ToObject<List<string>>().Count;

      for (int i = 0; i < length * 2; i += 2)
      {
        foreach (string key in multi_row_datagrid_keys)
        {
          if (selectedPanelData[key] is Newtonsoft.Json.Linq.JArray)
          {
            List<string> values = ((Newtonsoft.Json.Linq.JArray)selectedPanelData[key]).ToObject<List<string>>();

            if (i < values.Count)
            {
              string currentValue = values[i];
              string nextValue = i + 1 < values.Count ? values[i + 1] : null;

              if (key.Contains("description") && currentValue == "SPACE")
              {
                continue; // skip this iteration if the value is "SPACE" for descriptions
              }

              if (key.Contains("phase") && currentValue == "0")
              {
                continue; // skip this iteration if the value is "0" for phases
              }

              if (nextValue != null)
              {
                if (key.Contains("phase") && nextValue != "0")
                {
                  currentValue = $"{currentValue},{nextValue}";
                }
                else if (key.Contains("description") && nextValue != "SPACE")
                {
                  currentValue = $"{currentValue},{nextValue}";
                }
                else if (key.Contains("circuit"))
                {
                  currentValue = currentValue.Replace("A", "");
                }
                else if (key.Contains("breaker") && nextValue != "")
                {
                  currentValue = $"{currentValue},{nextValue}";
                }
              }

              PANEL_GRID.Rows[i / 2].Cells[key].Value = currentValue;
            }
          }
          else
          {
            // Log or handle the unexpected type
            Console.WriteLine($"Warning: Value for key {key} is not a JArray");
          }
        }
      }
    }

    private void clear_current_modal_data()
    {
      // Clear TextBoxes
      BUS_RATING_INPUT.Text = string.Empty;
      MAIN_INPUT.Text = string.Empty;
      PANEL_LOCATION_INPUT.Text = string.Empty;
      PANEL_NAME_INPUT.Text = string.Empty;
      LARGEST_LCL_INPUT.Text = string.Empty;

      // Clear ComboBoxes
      STATUS_COMBOBOX.SelectedIndex = -1; // This will unselect all items
      MOUNTING_COMBOBOX.SelectedIndex = -1;
      WIRE_COMBOBOX.SelectedIndex = -1;
      PHASE_COMBOBOX.SelectedIndex = -1;
      PHASE_VOLTAGE_COMBOBOX.SelectedIndex = -1;
      LINE_VOLTAGE_COMBOBOX.SelectedIndex = -1;
      LOAD_PANEL_COMBOBOX.SelectedIndex = -1;

      // Clear DataGridViews
      PHASE_SUM_GRID.Rows[0].Cells[0].Value = "0";
      PHASE_SUM_GRID.Rows[0].Cells[1].Value = "0";
      TOTAL_VA_GRID.Rows[0].Cells[0].Value = "0";
      LCL_GRID.Rows[0].Cells[0].Value = "0";
      LCL_GRID.Rows[0].Cells[1].Value = "0";
      TOTAL_OTHER_LOAD_GRID.Rows[0].Cells[0].Value = "0";
      PANEL_LOAD_GRID.Rows[0].Cells[0].Value = "0";
      FEEDER_AMP_GRID.Rows[0].Cells[0].Value = "0";

      // Clear DataGridViews
      for (int i = 0; i < PANEL_GRID.Rows.Count; i++)
      {
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

    private void NEW_PANEL_BUTTON_Click(object sender, EventArgs e)
    {
      // prompt the user if they would like to save the current panel
      DialogResult result = MessageBox.Show("Would you like to save the current panel?", "Confirm", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
      if (result == DialogResult.Yes)
      {
        save_panel_data();
      }
      else if (result == DialogResult.Cancel)
      {
        return;
      }

      clear_current_modal_data();

      // Set default values
      Set_Default_Form_Values();

      // remove all rows from PANEL_GRID
      PANEL_GRID.Rows.Clear();

      // add 21 rows to PANEL_GRID
      for (int i = 0; i < 21; i++)
      {
        PANEL_GRID.Rows.Add();
      }
    }
  }
}