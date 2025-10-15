namespace ElectricalCommands.SingleLine {
  partial class SingleLineDialogWindow {
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing) {
      if (disposing && (components != null)) {
        components.Dispose();
      }
      base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent() {
      this.components = new System.ComponentModel.Container();
      this.SingleLineTreeView = new System.Windows.Forms.TreeView();
      this.InfoGroupBox = new System.Windows.Forms.GroupBox();
      this.InfoTextBox = new System.Windows.Forms.TextBox();
      this.menuStrip1 = new System.Windows.Forms.MenuStrip();
      this.actionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.generateToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.placeEquipmentToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.refreshToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.InfoGroupBox.SuspendLayout();
      this.menuStrip1.SuspendLayout();
      this.SuspendLayout();
      // 
      // SingleLineTreeView
      // 
      this.SingleLineTreeView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.SingleLineTreeView.Location = new System.Drawing.Point(12, 27);
      this.SingleLineTreeView.Name = "SingleLineTreeView";
      this.SingleLineTreeView.Size = new System.Drawing.Size(802, 722);
      this.SingleLineTreeView.TabIndex = 0;
      this.SingleLineTreeView.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.TreeView_OnNodeMouseClick);
      // 
      // InfoGroupBox
      // 
      this.InfoGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.InfoGroupBox.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
      this.InfoGroupBox.Controls.Add(this.InfoTextBox);
      this.InfoGroupBox.Location = new System.Drawing.Point(820, 4);
      this.InfoGroupBox.Name = "InfoGroupBox";
      this.InfoGroupBox.Size = new System.Drawing.Size(352, 745);
      this.InfoGroupBox.TabIndex = 1;
      this.InfoGroupBox.TabStop = false;
      this.InfoGroupBox.Text = "groupBox1";
      // 
      // InfoTextBox
      // 
      this.InfoTextBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
      this.InfoTextBox.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.InfoTextBox.ForeColor = System.Drawing.SystemColors.MenuText;
      this.InfoTextBox.Location = new System.Drawing.Point(6, 19);
      this.InfoTextBox.Multiline = true;
      this.InfoTextBox.Name = "InfoTextBox";
      this.InfoTextBox.ReadOnly = true;
      this.InfoTextBox.Size = new System.Drawing.Size(340, 720);
      this.InfoTextBox.TabIndex = 0;
      // 
      // menuStrip1
      // 
      this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.actionsToolStripMenuItem});
      this.menuStrip1.Location = new System.Drawing.Point(0, 0);
      this.menuStrip1.Name = "menuStrip1";
      this.menuStrip1.Size = new System.Drawing.Size(1184, 24);
      this.menuStrip1.TabIndex = 2;
      this.menuStrip1.Text = "menuStrip1";
      // 
      // actionsToolStripMenuItem
      // 
      this.actionsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.generateToolStripMenuItem,
            this.placeEquipmentToolStripMenuItem,
            this.refreshToolStripMenuItem});
      this.actionsToolStripMenuItem.Name = "actionsToolStripMenuItem";
      this.actionsToolStripMenuItem.Size = new System.Drawing.Size(59, 20);
      this.actionsToolStripMenuItem.Text = "Actions";
      // 
      // generateToolStripMenuItem
      // 
      this.generateToolStripMenuItem.Name = "generateToolStripMenuItem";
      this.generateToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
      this.generateToolStripMenuItem.Text = "Generate";
      this.generateToolStripMenuItem.Click += new System.EventHandler(this.GenerateButton_Click);
      // 
      // placeEquipmentToolStripMenuItem
      // 
      this.placeEquipmentToolStripMenuItem.Name = "placeEquipmentToolStripMenuItem";
      this.placeEquipmentToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
      this.placeEquipmentToolStripMenuItem.Text = "Place Equipment";
      this.placeEquipmentToolStripMenuItem.Click += new System.EventHandler(this.PlaceOnPlanButton_Click);
      // 
      // refreshToolStripMenuItem
      // 
      this.refreshToolStripMenuItem.Name = "refreshToolStripMenuItem";
      this.refreshToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
      this.refreshToolStripMenuItem.Text = "Refresh";
      this.refreshToolStripMenuItem.Click += new System.EventHandler(this.RefreshButton_Click);
      // 
      // SingleLineDialogWindow
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(1184, 761);
      this.Controls.Add(this.InfoGroupBox);
      this.Controls.Add(this.SingleLineTreeView);
      this.Controls.Add(this.menuStrip1);
      this.MainMenuStrip = this.menuStrip1;
      this.Name = "SingleLineDialogWindow";
      this.Text = "Single Line";
      this.Load += new System.EventHandler(this.SingleLineDialogWindow_Load);
      this.InfoGroupBox.ResumeLayout(false);
      this.InfoGroupBox.PerformLayout();
      this.menuStrip1.ResumeLayout(false);
      this.menuStrip1.PerformLayout();
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion

    private System.Windows.Forms.TreeView SingleLineTreeView;
    private System.Windows.Forms.GroupBox InfoGroupBox;
    private System.Windows.Forms.TextBox InfoTextBox;
    private System.Windows.Forms.MenuStrip menuStrip1;
    private System.Windows.Forms.ToolStripMenuItem actionsToolStripMenuItem;
    private System.Windows.Forms.ToolStripMenuItem generateToolStripMenuItem;
    private System.Windows.Forms.ToolStripMenuItem placeEquipmentToolStripMenuItem;
    private System.Windows.Forms.ToolStripMenuItem refreshToolStripMenuItem;
  }
}