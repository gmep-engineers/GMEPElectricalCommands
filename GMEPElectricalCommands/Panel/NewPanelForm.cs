using ElectricalCommands;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ElectricalCommands {

  public partial class NewPanelForm : Form {
    private MainForm _mainForm;

    public NewPanelForm(MainForm mainForm) {
      InitializeComponent();
      this.StartPosition = FormStartPosition.CenterParent;
      _mainForm = mainForm;
    }

    private void CREATEPANEL_Click(object sender, EventArgs e) {
      // get the state of the checkbox
      bool is3PH = CHECKBOX3PH.Checked;

      // get the value of the textbox
      string panelName = CREATEPANELNAME.Text;

      // check if the panel name already exists
      if (_mainForm.PanelNameExists(panelName)) {
        MessageBox.Show("Panel name already exists. Please choose another name.");
        return;
      }

      // check if the panel name is empty
      if (panelName == "") {
        MessageBox.Show("Panel name cannot be empty.");
        return;
      }

      // call a method on the main form
      if (_mainForm != null) {
        var userControl = _mainForm.CreateNewPanelTab(panelName, is3PH);
        userControl.AddListeners();
      }
    }
  }
}