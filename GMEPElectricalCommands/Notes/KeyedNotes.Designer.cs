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
            this.components = new System.ComponentModel.Container();
            this.TableTabControl = new System.Windows.Forms.TabControl();
            this.button1 = new System.Windows.Forms.Button();
            this.TabMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.deleteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.TabMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // TableTabControl
            // 
            this.TableTabControl.Location = new System.Drawing.Point(12, 12);
            this.TableTabControl.Name = "TableTabControl";
            this.TableTabControl.SelectedIndex = 0;
            this.TableTabControl.Size = new System.Drawing.Size(776, 403);
            this.TableTabControl.TabIndex = 1;
            this.TableTabControl.MouseUp += new System.Windows.Forms.MouseEventHandler(this.TableTabControl_MouseUp);
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
            // TabMenu
            // 
            this.TabMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.deleteToolStripMenuItem});
            this.TabMenu.Name = "TabMenu";
            this.TabMenu.Size = new System.Drawing.Size(181, 48);
            // 
            // deleteToolStripMenuItem
            // 
            this.deleteToolStripMenuItem.Name = "deleteToolStripMenuItem";
            this.deleteToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.deleteToolStripMenuItem.Text = "Delete";
            this.deleteToolStripMenuItem.Click += new System.EventHandler(this.deleteToolStripMenuItem_Click);
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
            this.TabMenu.ResumeLayout(false);
            this.ResumeLayout(false);

    }

    #endregion
    private System.Windows.Forms.TabControl TableTabControl;
    private System.Windows.Forms.Button button1;
    private System.Windows.Forms.ContextMenuStrip TabMenu;
    private System.Windows.Forms.ToolStripMenuItem deleteToolStripMenuItem;
  }
}