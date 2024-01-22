﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ElectricalCommands
{
  public partial class Help : Form
  {
    public Help()
    {
      InitializeComponent();

      this.Shown += Help_Shown;
    }

    private void Help_Shown(object sender, EventArgs e)
    {
      HELP_TEXTBOX.SelectionLength = 0;
    }
  }
}