using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ElectricalCommands.Notes
{
  public partial class KeyedNotes: Form
  {
    // This is a placeholder for the DataTable that will hold the keyed notes
    public KeyedNotes()
    { 
        InitializeComponent();
        this.Load += new EventHandler(TabControl_Load);
    }

    public void TabControl_Load(object sender, EventArgs e) {
      //Add All Existing Tabs
      //TableTabControl.TabPages.Add(new TabPage("MEOW"));
      //Add 'New' Tab
      AddNewTabButton();
      // Set the initial selected tab to the last one (the one before "ADD NEW" tab)
      TableTabControl.SelectedIndex = TableTabControl.TabCount - 2;
    }

    private void AddNewTabButton() {
      TabPage addNewTabPage = new TabPage("ADD NEW") {
        BackColor = Color.AliceBlue
      };
      TableTabControl.TabPages.Add(addNewTabPage);
      TableTabControl.SelectedIndexChanged += TableTabControl_SelectedIndexChanged;
    }

    private void TableTabControl_SelectedIndexChanged(object sender, EventArgs e) {
      if (TableTabControl.SelectedTab != null && TableTabControl.SelectedTab.Text == "ADD NEW") {
        // Add a new tab before the "ADD NEW" tab
        int newIndex = TableTabControl.TabCount;
        TabPage newTab = new TabPage($"Tab {newIndex}");
        newTab.Controls.Add(new NoteTableUserControl());
        TableTabControl.TabPages.Insert(TableTabControl.TabCount - 1, newTab);
        TableTabControl.SelectedTab = newTab;
      }
    }
  }
}
