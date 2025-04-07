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
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // TableTabControl
            // 
            this.TableTabControl.Location = new System.Drawing.Point(12, 12);
            this.TableTabControl.Name = "TableTabControl";
            this.TableTabControl.SelectedIndex = 0;
            this.TableTabControl.Size = new System.Drawing.Size(776, 403);
            this.TableTabControl.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(359, 421);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "SAVE";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Save_Click);
            // 
            // KeyedNotes
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.TableTabControl);
            this.Name = "KeyedNotes";
            this.Text = "KeyedNotes";
            this.ResumeLayout(false);

    }

    #endregion
    private System.Windows.Forms.TabControl TableTabControl;
    private System.Windows.Forms.Button button1;
  }
}