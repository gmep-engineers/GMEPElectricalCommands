namespace ElectricalCommands.PlanCheck {
  partial class PlanCheckDialogWindow {
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
      this.RunAllButton = new System.Windows.Forms.Button();
      this.PlanCheckListView = new System.Windows.Forms.ListView();
      this.SuspendLayout();
      // 
      // RunAllButton
      // 
      this.RunAllButton.Location = new System.Drawing.Point(12, 439);
      this.RunAllButton.Name = "RunAllButton";
      this.RunAllButton.Size = new System.Drawing.Size(111, 23);
      this.RunAllButton.TabIndex = 3;
      this.RunAllButton.Text = "Run All";
      this.RunAllButton.UseVisualStyleBackColor = true;
      this.RunAllButton.Click += new System.EventHandler(this.RunAllButton_Click);
      // 
      // PlanCheckListView
      // 
      this.PlanCheckListView.HideSelection = false;
      this.PlanCheckListView.Location = new System.Drawing.Point(12, 12);
      this.PlanCheckListView.Name = "PlanCheckListView";
      this.PlanCheckListView.Size = new System.Drawing.Size(776, 421);
      this.PlanCheckListView.TabIndex = 4;
      this.PlanCheckListView.UseCompatibleStateImageBehavior = false;
      // 
      // PlanCheckDialogWindow
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(800, 469);
      this.Controls.Add(this.PlanCheckListView);
      this.Controls.Add(this.RunAllButton);
      this.Name = "PlanCheckDialogWindow";
      this.Text = "Electrical Plan Check";
      this.ResumeLayout(false);

    }

    #endregion
    private System.Windows.Forms.Button RunAllButton;
    private System.Windows.Forms.ListView PlanCheckListView;
  }
}