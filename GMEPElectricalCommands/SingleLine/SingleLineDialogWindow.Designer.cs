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
      this.mainMenu1 = new System.Windows.Forms.MainMenu(this.components);
      this.menuItem1 = new System.Windows.Forms.MenuItem();
      this.GenerateSldButton = new System.Windows.Forms.MenuItem();
      this.PlaceOnPlanButton = new System.Windows.Forms.MenuItem();
      this.menuItem2 = new System.Windows.Forms.MenuItem();
      this.InfoGroupBox = new System.Windows.Forms.GroupBox();
      this.InfoTextBox = new System.Windows.Forms.TextBox();
      this.InfoGroupBox.SuspendLayout();
      this.SuspendLayout();
      // 
      // SingleLineTreeView
      // 
      this.SingleLineTreeView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.SingleLineTreeView.Location = new System.Drawing.Point(12, 12);
      this.SingleLineTreeView.Name = "SingleLineTreeView";
      this.SingleLineTreeView.Size = new System.Drawing.Size(802, 737);
      this.SingleLineTreeView.TabIndex = 0;
      this.SingleLineTreeView.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.TreeView_OnNodeMouseClick);
      // 
      // mainMenu1
      // 
      this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem1});
      // 
      // menuItem1
      // 
      this.menuItem1.Index = 0;
      this.menuItem1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.GenerateSldButton,
            this.PlaceOnPlanButton,
            this.menuItem2});
      this.menuItem1.Text = "Actions";
      // 
      // GenerateSldButton
      // 
      this.GenerateSldButton.Index = 0;
      this.GenerateSldButton.Text = "Generate";
      this.GenerateSldButton.Click += new System.EventHandler(this.GenerateButton_Click);
      // 
      // PlaceOnPlanButton
      // 
      this.PlaceOnPlanButton.Index = 1;
      this.PlaceOnPlanButton.Text = "Place equipment";
      this.PlaceOnPlanButton.Click += new System.EventHandler(this.PlaceOnPlanButton_Click);
      // 
      // menuItem2
      // 
      this.menuItem2.Index = 2;
      this.menuItem2.Text = "Refresh";
      this.menuItem2.Click += new System.EventHandler(this.RefreshButton_Click);
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
      // SingleLineDialogWindow
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(1184, 761);
      this.Controls.Add(this.InfoGroupBox);
      this.Controls.Add(this.SingleLineTreeView);
      this.Menu = this.mainMenu1;
      this.Name = "SingleLineDialogWindow";
      this.Text = "Single Line";
      this.Load += new System.EventHandler(this.SingleLineDialogWindow_Load);
      this.InfoGroupBox.ResumeLayout(false);
      this.InfoGroupBox.PerformLayout();
      this.ResumeLayout(false);

    }

    #endregion

    private System.Windows.Forms.TreeView SingleLineTreeView;
    private System.Windows.Forms.MainMenu mainMenu1;
    private System.Windows.Forms.MenuItem menuItem1;
    private System.Windows.Forms.MenuItem GenerateSldButton;
    private System.Windows.Forms.GroupBox InfoGroupBox;
    private System.Windows.Forms.TextBox InfoTextBox;
    private System.Windows.Forms.MenuItem PlaceOnPlanButton;
    private System.Windows.Forms.MenuItem menuItem2;
  }
}