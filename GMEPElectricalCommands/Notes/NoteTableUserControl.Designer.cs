namespace ElectricalCommands.Notes {
  partial class NoteTableUserControl {
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

    #region Component Designer generated code

    /// <summary> 
    /// Required method for Designer support - do not modify 
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent() {
            this.components = new System.ComponentModel.Container();
            this.TableGridView = new System.Windows.Forms.DataGridView();
            this.gridMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.deleteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.placeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.TableGridView)).BeginInit();
            this.gridMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // TableGridView
            // 
            this.TableGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.TableGridView.Location = new System.Drawing.Point(3, 3);
            this.TableGridView.Name = "TableGridView";
            this.TableGridView.Size = new System.Drawing.Size(764, 414);
            this.TableGridView.TabIndex = 0;
            this.TableGridView.CellMouseUp += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.Grid_MouseUp);
            // 
            // gridMenu
            // 
            this.gridMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.deleteToolStripMenuItem,
            this.placeToolStripMenuItem});
            this.gridMenu.Name = "gridMenuStrip";
            this.gridMenu.Size = new System.Drawing.Size(181, 70);
            // 
            // deleteToolStripMenuItem
            // 
            this.deleteToolStripMenuItem.Name = "deleteToolStripMenuItem";
            this.deleteToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.deleteToolStripMenuItem.Text = "Delete";
            this.deleteToolStripMenuItem.Click += new System.EventHandler(this.deleteToolStripMenuItem_Click);
            // 
            // placeToolStripMenuItem
            // 
            this.placeToolStripMenuItem.Name = "placeToolStripMenuItem";
            this.placeToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.placeToolStripMenuItem.Text = "Place";
            this.placeToolStripMenuItem.Click += new System.EventHandler(this.placeToolStripMenuItem_Click);
            // 
            // NoteTableUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.TableGridView);
            this.Location = new System.Drawing.Point(0, 1);
            this.Name = "NoteTableUserControl";
            this.Size = new System.Drawing.Size(770, 420);
            this.Load += new System.EventHandler(this.NoteTableUserControl_Load);
            ((System.ComponentModel.ISupportInitialize)(this.TableGridView)).EndInit();
            this.gridMenu.ResumeLayout(false);
            this.ResumeLayout(false);

    }

    #endregion

    private System.Windows.Forms.DataGridView TableGridView;
    private System.Windows.Forms.ContextMenuStrip gridMenu;
    private System.Windows.Forms.ToolStripMenuItem deleteToolStripMenuItem;
    private System.Windows.Forms.ToolStripMenuItem placeToolStripMenuItem;
  }
}
