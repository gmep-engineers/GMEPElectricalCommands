namespace ElectricalCommands.Equipment {
  partial class EquipmentDialogWindow {
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
      this.filtersGroupBox = new System.Windows.Forms.GroupBox();
      this.filterClearButton = new System.Windows.Forms.Button();
      this.filterCategoryComboBox = new System.Windows.Forms.ComboBox();
      this.filterCategoryLabel = new System.Windows.Forms.Label();
      this.filterEquipNoTextBox = new System.Windows.Forms.TextBox();
      this.filterEquipNoLabel = new System.Windows.Forms.Label();
      this.filterPhaseLabel = new System.Windows.Forms.Label();
      this.filterPhaseComboBox = new System.Windows.Forms.ComboBox();
      this.filterVoltageComboBox = new System.Windows.Forms.ComboBox();
      this.filterVoltageLabel = new System.Windows.Forms.Label();
      this.filterPanelLabel = new System.Windows.Forms.Label();
      this.filterPanelComboBox = new System.Windows.Forms.ComboBox();
      this.equipmentGroupBox = new System.Windows.Forms.GroupBox();
      this.equipmentListView = new System.Windows.Forms.ListView();
      this.feedersGroupBox = new System.Windows.Forms.GroupBox();
      this.feederListView = new System.Windows.Forms.ListView();
      this.placeSelectedButton = new System.Windows.Forms.Button();
      this.filtersGroupBox.SuspendLayout();
      this.equipmentGroupBox.SuspendLayout();
      this.feedersGroupBox.SuspendLayout();
      this.SuspendLayout();
      // 
      // filtersGroupBox
      // 
      this.filtersGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.filtersGroupBox.Controls.Add(this.filterClearButton);
      this.filtersGroupBox.Controls.Add(this.filterCategoryComboBox);
      this.filtersGroupBox.Controls.Add(this.filterCategoryLabel);
      this.filtersGroupBox.Controls.Add(this.filterEquipNoTextBox);
      this.filtersGroupBox.Controls.Add(this.filterEquipNoLabel);
      this.filtersGroupBox.Controls.Add(this.filterPhaseLabel);
      this.filtersGroupBox.Controls.Add(this.filterPhaseComboBox);
      this.filtersGroupBox.Controls.Add(this.filterVoltageComboBox);
      this.filtersGroupBox.Controls.Add(this.filterVoltageLabel);
      this.filtersGroupBox.Controls.Add(this.filterPanelLabel);
      this.filtersGroupBox.Controls.Add(this.filterPanelComboBox);
      this.filtersGroupBox.Location = new System.Drawing.Point(12, 251);
      this.filtersGroupBox.Name = "filtersGroupBox";
      this.filtersGroupBox.Size = new System.Drawing.Size(860, 56);
      this.filtersGroupBox.TabIndex = 0;
      this.filtersGroupBox.TabStop = false;
      this.filtersGroupBox.Text = "Filters";
      // 
      // filterClearButton
      // 
      this.filterClearButton.Location = new System.Drawing.Point(757, 20);
      this.filterClearButton.Name = "filterClearButton";
      this.filterClearButton.Size = new System.Drawing.Size(75, 23);
      this.filterClearButton.TabIndex = 10;
      this.filterClearButton.Text = "Clear";
      this.filterClearButton.UseVisualStyleBackColor = true;
      this.filterClearButton.Click += new System.EventHandler(this.filterClearButton_Click);
      // 
      // filterCategoryComboBox
      // 
      this.filterCategoryComboBox.FormattingEnabled = true;
      this.filterCategoryComboBox.Location = new System.Drawing.Point(605, 22);
      this.filterCategoryComboBox.Name = "filterCategoryComboBox";
      this.filterCategoryComboBox.Size = new System.Drawing.Size(121, 21);
      this.filterCategoryComboBox.TabIndex = 9;
      // 
      // filterCategoryLabel
      // 
      this.filterCategoryLabel.AutoSize = true;
      this.filterCategoryLabel.Location = new System.Drawing.Point(550, 26);
      this.filterCategoryLabel.Name = "filterCategoryLabel";
      this.filterCategoryLabel.Size = new System.Drawing.Size(49, 13);
      this.filterCategoryLabel.TabIndex = 8;
      this.filterCategoryLabel.Text = "Category";
      // 
      // filterEquipNoTextBox
      // 
      this.filterEquipNoTextBox.Location = new System.Drawing.Point(427, 23);
      this.filterEquipNoTextBox.Name = "filterEquipNoTextBox";
      this.filterEquipNoTextBox.Size = new System.Drawing.Size(100, 20);
      this.filterEquipNoTextBox.TabIndex = 7;
      // 
      // filterEquipNoLabel
      // 
      this.filterEquipNoLabel.AutoSize = true;
      this.filterEquipNoLabel.Location = new System.Drawing.Point(376, 26);
      this.filterEquipNoLabel.Name = "filterEquipNoLabel";
      this.filterEquipNoLabel.Size = new System.Drawing.Size(44, 13);
      this.filterEquipNoLabel.TabIndex = 6;
      this.filterEquipNoLabel.Text = "Equip #";
      // 
      // filterPhaseLabel
      // 
      this.filterPhaseLabel.AutoSize = true;
      this.filterPhaseLabel.Location = new System.Drawing.Point(266, 26);
      this.filterPhaseLabel.Name = "filterPhaseLabel";
      this.filterPhaseLabel.Size = new System.Drawing.Size(37, 13);
      this.filterPhaseLabel.TabIndex = 5;
      this.filterPhaseLabel.Text = "Phase";
      // 
      // filterPhaseComboBox
      // 
      this.filterPhaseComboBox.FormattingEnabled = true;
      this.filterPhaseComboBox.Items.AddRange(new object[] {
            "1",
            "3"});
      this.filterPhaseComboBox.Location = new System.Drawing.Point(309, 23);
      this.filterPhaseComboBox.Name = "filterPhaseComboBox";
      this.filterPhaseComboBox.Size = new System.Drawing.Size(50, 21);
      this.filterPhaseComboBox.TabIndex = 4;
      // 
      // filterVoltageComboBox
      // 
      this.filterVoltageComboBox.FormattingEnabled = true;
      this.filterVoltageComboBox.Items.AddRange(new object[] {
            "115",
            "120",
            "208",
            "230",
            "240",
            "460",
            "480"});
      this.filterVoltageComboBox.Location = new System.Drawing.Point(175, 23);
      this.filterVoltageComboBox.Name = "filterVoltageComboBox";
      this.filterVoltageComboBox.Size = new System.Drawing.Size(72, 21);
      this.filterVoltageComboBox.TabIndex = 3;
      this.filterVoltageComboBox.SelectedIndexChanged += new System.EventHandler(this.FilterVoltageComboBox_SelectedIndexChanged);
      // 
      // filterVoltageLabel
      // 
      this.filterVoltageLabel.AutoSize = true;
      this.filterVoltageLabel.Location = new System.Drawing.Point(126, 26);
      this.filterVoltageLabel.Name = "filterVoltageLabel";
      this.filterVoltageLabel.Size = new System.Drawing.Size(43, 13);
      this.filterVoltageLabel.TabIndex = 2;
      this.filterVoltageLabel.Text = "Voltage";
      // 
      // filterPanelLabel
      // 
      this.filterPanelLabel.AutoSize = true;
      this.filterPanelLabel.Location = new System.Drawing.Point(6, 26);
      this.filterPanelLabel.Name = "filterPanelLabel";
      this.filterPanelLabel.Size = new System.Drawing.Size(34, 13);
      this.filterPanelLabel.TabIndex = 1;
      this.filterPanelLabel.Text = "Panel";
      // 
      // filterPanelComboBox
      // 
      this.filterPanelComboBox.FormattingEnabled = true;
      this.filterPanelComboBox.Location = new System.Drawing.Point(46, 23);
      this.filterPanelComboBox.Name = "filterPanelComboBox";
      this.filterPanelComboBox.Size = new System.Drawing.Size(63, 21);
      this.filterPanelComboBox.TabIndex = 0;
      this.filterPanelComboBox.SelectedIndexChanged += new System.EventHandler(this.FilterPanelComboBox_SelectedIndexChanged);
      // 
      // equipmentGroupBox
      // 
      this.equipmentGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.equipmentGroupBox.Controls.Add(this.equipmentListView);
      this.equipmentGroupBox.Location = new System.Drawing.Point(12, 313);
      this.equipmentGroupBox.Name = "equipmentGroupBox";
      this.equipmentGroupBox.Size = new System.Drawing.Size(860, 502);
      this.equipmentGroupBox.TabIndex = 1;
      this.equipmentGroupBox.TabStop = false;
      this.equipmentGroupBox.Text = "Equipment";
      // 
      // equipmentListView
      // 
      this.equipmentListView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.equipmentListView.HideSelection = false;
      this.equipmentListView.Location = new System.Drawing.Point(7, 20);
      this.equipmentListView.Name = "equipmentListView";
      this.equipmentListView.Size = new System.Drawing.Size(847, 476);
      this.equipmentListView.TabIndex = 0;
      this.equipmentListView.UseCompatibleStateImageBehavior = false;
      this.equipmentListView.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.EquipmentListView_MouseDoubleClick);
      // 
      // feedersGroupBox
      // 
      this.feedersGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.feedersGroupBox.Controls.Add(this.feederListView);
      this.feedersGroupBox.Location = new System.Drawing.Point(12, 13);
      this.feedersGroupBox.Name = "feedersGroupBox";
      this.feedersGroupBox.Size = new System.Drawing.Size(860, 232);
      this.feedersGroupBox.TabIndex = 2;
      this.feedersGroupBox.TabStop = false;
      this.feedersGroupBox.Text = "Feeders";
      // 
      // feederListView
      // 
      this.feederListView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.feederListView.HideSelection = false;
      this.feederListView.Location = new System.Drawing.Point(6, 19);
      this.feederListView.Name = "feederListView";
      this.feederListView.Size = new System.Drawing.Size(848, 207);
      this.feederListView.TabIndex = 0;
      this.feederListView.UseCompatibleStateImageBehavior = false;
      // 
      // placeSelectedButton
      // 
      this.placeSelectedButton.Location = new System.Drawing.Point(21, 826);
      this.placeSelectedButton.Name = "placeSelectedButton";
      this.placeSelectedButton.Size = new System.Drawing.Size(124, 23);
      this.placeSelectedButton.TabIndex = 3;
      this.placeSelectedButton.Text = "Place Selected";
      this.placeSelectedButton.UseVisualStyleBackColor = true;
      this.placeSelectedButton.Click += new System.EventHandler(this.PlaceSelectedButton_Click);
      // 
      // EquipmentDialogWindow
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(884, 861);
      this.Controls.Add(this.placeSelectedButton);
      this.Controls.Add(this.feedersGroupBox);
      this.Controls.Add(this.equipmentGroupBox);
      this.Controls.Add(this.filtersGroupBox);
      this.Name = "EquipmentDialogWindow";
      this.Text = "EquipmentDialogWindow";
      this.filtersGroupBox.ResumeLayout(false);
      this.filtersGroupBox.PerformLayout();
      this.equipmentGroupBox.ResumeLayout(false);
      this.feedersGroupBox.ResumeLayout(false);
      this.ResumeLayout(false);

    }

    #endregion

    private System.Windows.Forms.GroupBox filtersGroupBox;
    private System.Windows.Forms.GroupBox equipmentGroupBox;
    private System.Windows.Forms.Label filterVoltageLabel;
    private System.Windows.Forms.Label filterPanelLabel;
    private System.Windows.Forms.ComboBox filterPanelComboBox;
    private System.Windows.Forms.ComboBox filterVoltageComboBox;
    private System.Windows.Forms.TextBox filterEquipNoTextBox;
    private System.Windows.Forms.Label filterEquipNoLabel;
    private System.Windows.Forms.Label filterPhaseLabel;
    private System.Windows.Forms.ComboBox filterPhaseComboBox;
    private System.Windows.Forms.Label filterCategoryLabel;
    private System.Windows.Forms.ComboBox filterCategoryComboBox;
    private System.Windows.Forms.ListView equipmentListView;
    private System.Windows.Forms.Button filterClearButton;
    private System.Windows.Forms.GroupBox feedersGroupBox;
    private System.Windows.Forms.ListView feederListView;
    private System.Windows.Forms.Button placeSelectedButton;
  }
}