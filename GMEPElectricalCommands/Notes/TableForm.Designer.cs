namespace ElectricalCommands.Notes {
  partial class TableForm {
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
            this.button1 = new System.Windows.Forms.Button();
            this.SheetName = new System.Windows.Forms.RichTextBox();
            this.TableType = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.TableSheet = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(146, 246);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Enter";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // SheetName
            // 
            this.SheetName.Location = new System.Drawing.Point(91, 36);
            this.SheetName.Name = "SheetName";
            this.SheetName.Size = new System.Drawing.Size(189, 34);
            this.SheetName.TabIndex = 1;
            this.SheetName.Text = "";
            this.SheetName.TextChanged += new System.EventHandler(this.SheetName_TextChanged);
            // 
            // TableType
            // 
            this.TableType.FormattingEnabled = true;
            this.TableType.Items.AddRange(new object[] {
            "Single Line",
            "Electrical Power",
            "Electrical Lighting",
            "Electrical Roof"});
            this.TableType.Location = new System.Drawing.Point(91, 167);
            this.TableType.Name = "TableType";
            this.TableType.Size = new System.Drawing.Size(189, 21);
            this.TableType.TabIndex = 4;
            this.TableType.SelectedIndexChanged += new System.EventHandler(this.TableType_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(156, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Table Name";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(173, 89);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Sheet";
            // 
            // TableSheet
            // 
            this.TableSheet.FormattingEnabled = true;
            this.TableSheet.Location = new System.Drawing.Point(91, 105);
            this.TableSheet.Name = "TableSheet";
            this.TableSheet.Size = new System.Drawing.Size(189, 21);
            this.TableSheet.TabIndex = 7;
            this.TableSheet.SelectedIndexChanged += new System.EventHandler(this.TableSheet_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(173, 151);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(31, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Type";
            // 
            // TableForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(393, 281);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.TableSheet);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.TableType);
            this.Controls.Add(this.SheetName);
            this.Controls.Add(this.button1);
            this.Name = "TableForm";
            this.Text = "TableForm";
            this.ResumeLayout(false);
            this.PerformLayout();

    }

    #endregion

    private System.Windows.Forms.Button button1;
    private System.Windows.Forms.RichTextBox SheetName;
    private System.Windows.Forms.ComboBox TableType;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.Label label2;
    private System.Windows.Forms.ComboBox TableSheet;
    private System.Windows.Forms.Label label3;
  }
}