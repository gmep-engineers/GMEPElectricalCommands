namespace ElectricalCommands.Notes {
  partial class KeyedNotes {
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
            this.TableTabControl = new System.Windows.Forms.TabControl();
            this.SuspendLayout();
            // 
            // TableTabControl
            // 
            this.TableTabControl.Location = new System.Drawing.Point(12, 12);
            this.TableTabControl.Name = "TableTabControl";
            this.TableTabControl.SelectedIndex = 0;
            this.TableTabControl.Size = new System.Drawing.Size(776, 426);
            this.TableTabControl.TabIndex = 1;
            // 
            // KeyedNotes
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.TableTabControl);
            this.Name = "KeyedNotes";
            this.Text = "KeyedNotes";
            this.ResumeLayout(false);

    }

    #endregion
    private System.Windows.Forms.TabControl TableTabControl;
  }
}