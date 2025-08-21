namespace ElectricalCommands.Lighting {
  partial class LightingDialogWindow {
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
      this.LightingSignageGroupBox = new System.Windows.Forms.GroupBox();
      this.LightingSignageListView = new System.Windows.Forms.ListView();
      this.ControlsGroupBox = new System.Windows.Forms.GroupBox();
      this.LightingControlsListView = new System.Windows.Forms.ListView();
      this.LightingFixturesGroupBox = new System.Windows.Forms.GroupBox();
      this.LightingFixturesListView = new System.Windows.Forms.ListView();
      this.CreateLightingFixtureScheduleButton = new System.Windows.Forms.Button();
      this.PlaceSelectedFixturesButton = new System.Windows.Forms.Button();
      this.RefreshButton = new System.Windows.Forms.Button();
      this.PlaceSelectedControlsButton = new System.Windows.Forms.Button();
      this.PlaceSelectedSignageButton = new System.Windows.Forms.Button();
      this.LightingSignageGroupBox.SuspendLayout();
      this.ControlsGroupBox.SuspendLayout();
      this.LightingFixturesGroupBox.SuspendLayout();
      this.SuspendLayout();
      // 
      // SignageGroupBox
      // 
      this.LightingSignageGroupBox.Controls.Add(this.LightingSignageListView);
      this.LightingSignageGroupBox.Location = new System.Drawing.Point(12, 12);
      this.LightingSignageGroupBox.Name = "LightingSignageGroupBox";
      this.LightingSignageGroupBox.Size = new System.Drawing.Size(508, 225);
      this.LightingSignageGroupBox.TabIndex = 0;
      this.LightingSignageGroupBox.TabStop = false;
      this.LightingSignageGroupBox.Text = "Signage";
      // 
      // SignageListView
      // 
      this.LightingSignageListView.HideSelection = false;
      this.LightingSignageListView.Location = new System.Drawing.Point(7, 20);
      this.LightingSignageListView.Name = "LightingSignageListView";
      this.LightingSignageListView.Size = new System.Drawing.Size(495, 199);
      this.LightingSignageListView.TabIndex = 0;
      this.LightingSignageListView.UseCompatibleStateImageBehavior = false;
      // 
      // ControlsGroupBox
      // 
      this.ControlsGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
      this.ControlsGroupBox.Controls.Add(this.LightingControlsListView);
      this.ControlsGroupBox.Location = new System.Drawing.Point(526, 12);
      this.ControlsGroupBox.Name = "ControlsGroupBox";
      this.ControlsGroupBox.Size = new System.Drawing.Size(346, 225);
      this.ControlsGroupBox.TabIndex = 1;
      this.ControlsGroupBox.TabStop = false;
      this.ControlsGroupBox.Text = "Controls";
      // 
      // LightingControlsListView
      // 
      this.LightingControlsListView.HideSelection = false;
      this.LightingControlsListView.Location = new System.Drawing.Point(7, 20);
      this.LightingControlsListView.Name = "LightingControlsListView";
      this.LightingControlsListView.Size = new System.Drawing.Size(333, 199);
      this.LightingControlsListView.TabIndex = 0;
      this.LightingControlsListView.UseCompatibleStateImageBehavior = false;
      // 
      // LightingFixturesGroupBox
      // 
      this.LightingFixturesGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.LightingFixturesGroupBox.Controls.Add(this.LightingFixturesListView);
      this.LightingFixturesGroupBox.Location = new System.Drawing.Point(12, 243);
      this.LightingFixturesGroupBox.Name = "LightingFixturesGroupBox";
      this.LightingFixturesGroupBox.Size = new System.Drawing.Size(860, 177);
      this.LightingFixturesGroupBox.TabIndex = 2;
      this.LightingFixturesGroupBox.TabStop = false;
      this.LightingFixturesGroupBox.Text = "Lighting Fixtures";
      // 
      // LightingFixturesListView
      // 
      this.LightingFixturesListView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.LightingFixturesListView.HideSelection = false;
      this.LightingFixturesListView.Location = new System.Drawing.Point(7, 20);
      this.LightingFixturesListView.Name = "LightingFixturesListView";
      this.LightingFixturesListView.Size = new System.Drawing.Size(847, 151);
      this.LightingFixturesListView.TabIndex = 0;
      this.LightingFixturesListView.UseCompatibleStateImageBehavior = false;
      // 
      // CreateLightingFixtureScheduleButton
      // 
      this.CreateLightingFixtureScheduleButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
      this.CreateLightingFixtureScheduleButton.Location = new System.Drawing.Point(683, 426);
      this.CreateLightingFixtureScheduleButton.Name = "CreateLightingFixtureScheduleButton";
      this.CreateLightingFixtureScheduleButton.Size = new System.Drawing.Size(189, 23);
      this.CreateLightingFixtureScheduleButton.TabIndex = 3;
      this.CreateLightingFixtureScheduleButton.Text = "Create Lighting Fixture Schedule";
      this.CreateLightingFixtureScheduleButton.UseVisualStyleBackColor = true;
      this.CreateLightingFixtureScheduleButton.Click += new System.EventHandler(this.CreateLightingFixtureScheduleButton_Click);
      // 
      // RefreshButton
      // 
      this.RefreshButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
      this.RefreshButton.Location = new System.Drawing.Point(602, 426);
      this.RefreshButton.Name = "RefreshButton";
      this.RefreshButton.Size = new System.Drawing.Size(75, 23);
      this.RefreshButton.TabIndex = 5;
      this.RefreshButton.Text = "Refresh";
      this.RefreshButton.UseVisualStyleBackColor = true;
      this.RefreshButton.Click += new System.EventHandler(this.RefreshButton_Click);
      // 
      // PlaceSelectedFixturesButton
      // 
      this.PlaceSelectedFixturesButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
      this.PlaceSelectedFixturesButton.Location = new System.Drawing.Point(19, 425);
      this.PlaceSelectedFixturesButton.Name = "PlaceSelectedFixturesButton";
      this.PlaceSelectedFixturesButton.Size = new System.Drawing.Size(149, 23);
      this.PlaceSelectedFixturesButton.TabIndex = 4;
      this.PlaceSelectedFixturesButton.Text = "Place Selected Fixtures";
      this.PlaceSelectedFixturesButton.UseVisualStyleBackColor = true;
      this.PlaceSelectedFixturesButton.Click += new System.EventHandler(this.PlaceFixture_Click);

      // 
      // PlaceSelectedControlsButton
      // 
      this.PlaceSelectedControlsButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
      this.PlaceSelectedControlsButton.Location = new System.Drawing.Point(183, 425);
      this.PlaceSelectedControlsButton.Name = "PlaceSelectedControlsButton";
      this.PlaceSelectedControlsButton.Size = new System.Drawing.Size(152, 23);
      this.PlaceSelectedControlsButton.TabIndex = 6;
      this.PlaceSelectedControlsButton.Text = "Place Selected Controls";
      this.PlaceSelectedControlsButton.UseVisualStyleBackColor = true;
      this.PlaceSelectedControlsButton.Click += new System.EventHandler(this.PlaceControl_Click);
      // 
      // PlaceSelectedSignageButton
      // 
      this.PlaceSelectedSignageButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
      this.PlaceSelectedSignageButton.Location = new System.Drawing.Point(350, 425);
      this.PlaceSelectedSignageButton.Name = "PlaceSelectedSignageButton";
      this.PlaceSelectedSignageButton.Size = new System.Drawing.Size(152, 23);
      this. PlaceSelectedSignageButton.TabIndex = 7;
      this.PlaceSelectedSignageButton.Text = "Place Selected Signage";
      this.PlaceSelectedSignageButton.UseVisualStyleBackColor = true;
      //this.PlaceSelectedControlsButton.Click += new System.EventHandler(this.PlaceControl_Click);
      // 
      // LightingDialogWindow
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(884, 461);
      this.Controls.Add(this.PlaceSelectedControlsButton);
      this.Controls.Add(this.RefreshButton);
      this.Controls.Add(this.PlaceSelectedFixturesButton);
      this.Controls.Add(this.PlaceSelectedSignageButton);
      this.Controls.Add(this.CreateLightingFixtureScheduleButton);
      this.Controls.Add(this.LightingFixturesGroupBox);
      this.Controls.Add(this.ControlsGroupBox);
      this.Controls.Add(this.LightingSignageGroupBox);
      this.MaximumSize = new System.Drawing.Size(900, 1200);
      this.MinimumSize = new System.Drawing.Size(900, 500);
      this.Name = "LightingDialogWindow";
      this.Text = "Lighting";
      this.LightingSignageGroupBox.ResumeLayout(false);
      this.ControlsGroupBox.ResumeLayout(false);
      this.LightingFixturesGroupBox.ResumeLayout(false);
      this.ResumeLayout(false);

    }

    #endregion

    private System.Windows.Forms.GroupBox LightingSignageGroupBox;
    private System.Windows.Forms.GroupBox ControlsGroupBox;
    private System.Windows.Forms.GroupBox LightingFixturesGroupBox;
    private System.Windows.Forms.ListView LightingSignageListView;
    private System.Windows.Forms.ListView LightingControlsListView;
    private System.Windows.Forms.ListView LightingFixturesListView;
    private System.Windows.Forms.Button CreateLightingFixtureScheduleButton;
    private System.Windows.Forms.Button PlaceSelectedFixturesButton;
    private System.Windows.Forms.Button RefreshButton;
    private System.Windows.Forms.Button PlaceSelectedControlsButton;
    private System.Windows.Forms.Button PlaceSelectedSignageButton;
  }
}