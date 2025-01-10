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
      this.ControlAreasGroupBox = new System.Windows.Forms.GroupBox();
      this.ControlsGroupBox = new System.Windows.Forms.GroupBox();
      this.LightingFixturesGroupBox = new System.Windows.Forms.GroupBox();
      this.ControlAreasListView = new System.Windows.Forms.ListView();
      this.ControlsListView = new System.Windows.Forms.ListView();
      this.LightingFixturesListView = new System.Windows.Forms.ListView();
      this.CreateLightingFixtureDiagramButton = new System.Windows.Forms.Button();
      this.ControlAreasGroupBox.SuspendLayout();
      this.ControlsGroupBox.SuspendLayout();
      this.LightingFixturesGroupBox.SuspendLayout();
      this.SuspendLayout();
      // 
      // ControlAreasGroupBox
      // 
      this.ControlAreasGroupBox.Controls.Add(this.ControlAreasListView);
      this.ControlAreasGroupBox.Location = new System.Drawing.Point(12, 12);
      this.ControlAreasGroupBox.Name = "ControlAreasGroupBox";
      this.ControlAreasGroupBox.Size = new System.Drawing.Size(508, 225);
      this.ControlAreasGroupBox.TabIndex = 0;
      this.ControlAreasGroupBox.TabStop = false;
      this.ControlAreasGroupBox.Text = "Control Areas";
      // 
      // ControlsGroupBox
      // 
      this.ControlsGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
      this.ControlsGroupBox.Controls.Add(this.ControlsListView);
      this.ControlsGroupBox.Location = new System.Drawing.Point(526, 12);
      this.ControlsGroupBox.Name = "ControlsGroupBox";
      this.ControlsGroupBox.Size = new System.Drawing.Size(346, 225);
      this.ControlsGroupBox.TabIndex = 1;
      this.ControlsGroupBox.TabStop = false;
      this.ControlsGroupBox.Text = "Controls";
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
      // ControlAreasListView
      // 
      this.ControlAreasListView.HideSelection = false;
      this.ControlAreasListView.Location = new System.Drawing.Point(7, 20);
      this.ControlAreasListView.Name = "ControlAreasListView";
      this.ControlAreasListView.Size = new System.Drawing.Size(495, 199);
      this.ControlAreasListView.TabIndex = 0;
      this.ControlAreasListView.UseCompatibleStateImageBehavior = false;
      // 
      // ControlsListView
      // 
      this.ControlsListView.HideSelection = false;
      this.ControlsListView.Location = new System.Drawing.Point(7, 20);
      this.ControlsListView.Name = "ControlsListView";
      this.ControlsListView.Size = new System.Drawing.Size(333, 199);
      this.ControlsListView.TabIndex = 0;
      this.ControlsListView.UseCompatibleStateImageBehavior = false;
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
      // CreateLightingFixtureDiagramButton
      // 
      this.CreateLightingFixtureDiagramButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
      this.CreateLightingFixtureDiagramButton.Location = new System.Drawing.Point(19, 426);
      this.CreateLightingFixtureDiagramButton.Name = "CreateLightingFixtureDiagramButton";
      this.CreateLightingFixtureDiagramButton.Size = new System.Drawing.Size(189, 23);
      this.CreateLightingFixtureDiagramButton.TabIndex = 3;
      this.CreateLightingFixtureDiagramButton.Text = "Create Lighting Fixture Diagram";
      this.CreateLightingFixtureDiagramButton.UseVisualStyleBackColor = true;
      // 
      // LightingDialogWindow
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(884, 461);
      this.Controls.Add(this.CreateLightingFixtureDiagramButton);
      this.Controls.Add(this.LightingFixturesGroupBox);
      this.Controls.Add(this.ControlsGroupBox);
      this.Controls.Add(this.ControlAreasGroupBox);
      this.MaximumSize = new System.Drawing.Size(900, 1200);
      this.MinimumSize = new System.Drawing.Size(900, 500);
      this.Name = "LightingDialogWindow";
      this.Text = "Lighting";
      this.ControlAreasGroupBox.ResumeLayout(false);
      this.ControlsGroupBox.ResumeLayout(false);
      this.LightingFixturesGroupBox.ResumeLayout(false);
      this.ResumeLayout(false);

    }

    #endregion

    private System.Windows.Forms.GroupBox ControlAreasGroupBox;
    private System.Windows.Forms.GroupBox ControlsGroupBox;
    private System.Windows.Forms.GroupBox LightingFixturesGroupBox;
    private System.Windows.Forms.ListView ControlAreasListView;
    private System.Windows.Forms.ListView ControlsListView;
    private System.Windows.Forms.ListView LightingFixturesListView;
    private System.Windows.Forms.Button CreateLightingFixtureDiagramButton;
  }
}