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
            this.TableGridView = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.TableGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // TableGridView
            // 
            this.TableGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.TableGridView.Location = new System.Drawing.Point(3, 3);
            this.TableGridView.Name = "TableGridView";
            this.TableGridView.Size = new System.Drawing.Size(764, 414);
            this.TableGridView.TabIndex = 0;
            this.TableGridView.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.TableGridView_CellContentDoubleClick);

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
            this.ResumeLayout(false);

    }

    #endregion

    private System.Windows.Forms.DataGridView TableGridView;
  }
}
