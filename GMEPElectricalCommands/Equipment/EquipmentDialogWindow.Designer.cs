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
      this.panelsGroupBox = new System.Windows.Forms.GroupBox();
      this.panelListView = new System.Windows.Forms.ListView();
      this.placeSelectedButton = new System.Windows.Forms.Button();
      this.placeAllButton = new System.Windows.Forms.Button();
      this.recalculateDistancesButton = new System.Windows.Forms.Button();
      this.transformerGroupBox = new System.Windows.Forms.GroupBox();
      this.listView1 = new System.Windows.Forms.ListView();
      this.filtersGroupBox.SuspendLayout();
      this.equipmentGroupBox.SuspendLayout();
      this.panelsGroupBox.SuspendLayout();
      this.transformerGroupBox.SuspendLayout();
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
      this.filterClearButton.Click += new System.EventHandler(this.FilterClearButton_Click);
      // 
      // filterCategoryComboBox
      // 
      this.filterCategoryComboBox.FormattingEnabled = true;
      this.filterCategoryComboBox.Items.AddRange(new object[] {
            "General",
            "Lighting",
            "Mechanical",
            "Plumbing"});
      this.filterCategoryComboBox.Location = new System.Drawing.Point(605, 22);
      this.filterCategoryComboBox.Name = "filterCategoryComboBox";
      this.filterCategoryComboBox.Size = new System.Drawing.Size(121, 21);
      this.filterCategoryComboBox.TabIndex = 9;
      this.filterCategoryComboBox.SelectedIndexChanged += new System.EventHandler(this.FilterCategoryComboBox_SelectedIndexChanged);
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
      this.filterEquipNoTextBox.KeyUp += new System.Windows.Forms.KeyEventHandler(this.FilterEquipNoTextBox_KeyUp);
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
      this.filterPhaseComboBox.SelectedIndexChanged += new System.EventHandler(this.FilterPhaseComboBox_SelectedIndexChanged);
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
      this.equipmentListView.Location = new System.Drawing.Point(6, 20);
      this.equipmentListView.Name = "equipmentListView";
      this.equipmentListView.Size = new System.Drawing.Size(848, 476);
      this.equipmentListView.TabIndex = 0;
      this.equipmentListView.UseCompatibleStateImageBehavior = false;
      this.equipmentListView.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.EquipmentListView_MouseDoubleClick);
      // 
      // panelsGroupBox
      // 
      this.panelsGroupBox.Controls.Add(this.panelListView);
      this.panelsGroupBox.Location = new System.Drawing.Point(12, 13);
      this.panelsGroupBox.Name = "panelsGroupBox";
      this.panelsGroupBox.Size = new System.Drawing.Size(420, 232);
      this.panelsGroupBox.TabIndex = 2;
      this.panelsGroupBox.TabStop = false;
      this.panelsGroupBox.Text = "Panels";
      // 
      // panelListView
      // 
      this.panelListView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.panelListView.HideSelection = false;
      this.panelListView.Location = new System.Drawing.Point(6, 19);
      this.panelListView.Name = "panelListView";
      this.panelListView.Size = new System.Drawing.Size(408, 207);
      this.panelListView.TabIndex = 0;
      this.panelListView.UseCompatibleStateImageBehavior = false;
      this.panelListView.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.PanelListView_MouseDoubleClick);
      // 
      // placeSelectedButton
      // 
      this.placeSelectedButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
      this.placeSelectedButton.Location = new System.Drawing.Point(21, 826);
      this.placeSelectedButton.Name = "placeSelectedButton";
      this.placeSelectedButton.Size = new System.Drawing.Size(100, 23);
      this.placeSelectedButton.TabIndex = 3;
      this.placeSelectedButton.Text = "Place Selected";
      this.placeSelectedButton.UseVisualStyleBackColor = true;
      this.placeSelectedButton.Click += new System.EventHandler(this.PlaceSelectedButton_Click);
      // 
      // placeAllButton
      // 
      this.placeAllButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
      this.placeAllButton.Location = new System.Drawing.Point(141, 826);
      this.placeAllButton.Name = "placeAllButton";
      this.placeAllButton.Size = new System.Drawing.Size(75, 23);
      this.placeAllButton.TabIndex = 4;
      this.placeAllButton.Text = "Place All";
      this.placeAllButton.UseVisualStyleBackColor = true;
      this.placeAllButton.Click += new System.EventHandler(this.PlaceAllButton_Click);
      // 
      // recalculateDistancesButton
      // 
      this.recalculateDistancesButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
      this.recalculateDistancesButton.Location = new System.Drawing.Point(728, 826);
      this.recalculateDistancesButton.Name = "recalculateDistancesButton";
      this.recalculateDistancesButton.Size = new System.Drawing.Size(138, 23);
      this.recalculateDistancesButton.TabIndex = 5;
      this.recalculateDistancesButton.Text = "Recalculate Distances";
      this.recalculateDistancesButton.UseVisualStyleBackColor = true;
      this.recalculateDistancesButton.Click += new System.EventHandler(this.RecalculateDistancesButton_Click);
      // 
      // transformerGroupBox
      // 
      this.transformerGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
      this.transformerGroupBox.Controls.Add(this.listView1);
      this.transformerGroupBox.Location = new System.Drawing.Point(439, 13);
      this.transformerGroupBox.Name = "transformerGroupBox";
      this.transformerGroupBox.Size = new System.Drawing.Size(433, 232);
      this.transformerGroupBox.TabIndex = 6;
      this.transformerGroupBox.TabStop = false;
      this.transformerGroupBox.Text = "Transformers";
      // 
      // listView1
      // 
      this.listView1.HideSelection = false;
      this.listView1.Location = new System.Drawing.Point(7, 19);
      this.listView1.Name = "listView1";
      this.listView1.Size = new System.Drawing.Size(420, 207);
      this.listView1.TabIndex = 0;
      this.listView1.UseCompatibleStateImageBehavior = false;
      // 
      // EquipmentDialogWindow
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(884, 861);
      this.Controls.Add(this.transformerGroupBox);
      this.Controls.Add(this.recalculateDistancesButton);
      this.Controls.Add(this.placeAllButton);
      this.Controls.Add(this.placeSelectedButton);
      this.Controls.Add(this.panelsGroupBox);
      this.Controls.Add(this.equipmentGroupBox);
      this.Controls.Add(this.filtersGroupBox);
      this.MaximumSize = new System.Drawing.Size(900, 1200);
      this.MinimumSize = new System.Drawing.Size(900, 500);
      this.Name = "EquipmentDialogWindow";
      this.Text = "EquipmentDialogWindow";
      this.filtersGroupBox.ResumeLayout(false);
      this.filtersGroupBox.PerformLayout();
      this.equipmentGroupBox.ResumeLayout(false);
      this.panelsGroupBox.ResumeLayout(false);
      this.transformerGroupBox.ResumeLayout(false);
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
    private System.Windows.Forms.GroupBox panelsGroupBox;
    private System.Windows.Forms.ListView panelListView;
    private System.Windows.Forms.Button placeSelectedButton;
    private System.Windows.Forms.Button placeAllButton;
    private System.Windows.Forms.Button recalculateDistancesButton;
    private System.Windows.Forms.GroupBox transformerGroupBox;
    private System.Windows.Forms.ListView listView1;
  }
}