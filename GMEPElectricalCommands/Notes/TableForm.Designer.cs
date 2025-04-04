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
            this.TableName = new System.Windows.Forms.RichTextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.TableSheet = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(146, 229);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Enter";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // TableName
            // 
            this.TableName.Location = new System.Drawing.Point(91, 82);
            this.TableName.Name = "TableName";
            this.TableName.Size = new System.Drawing.Size(189, 34);
            this.TableName.TabIndex = 1;
            this.TableName.Text = "";
            this.TableName.TextChanged += new System.EventHandler(this.TableName_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(156, 57);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Table Name";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(172, 143);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(35, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Sheet";
            // 
            // TableSheet
            // 
            this.TableSheet.FormattingEnabled = true;
            this.TableSheet.Location = new System.Drawing.Point(91, 168);
            this.TableSheet.Name = "TableSheet";
            this.TableSheet.Size = new System.Drawing.Size(189, 21);
            this.TableSheet.TabIndex = 7;
            this.TableSheet.SelectedIndexChanged += new System.EventHandler(this.TableSheet_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.SystemColors.Control;
            this.label3.Font = new System.Drawing.Font("Microsoft YaHei UI", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label3.Location = new System.Drawing.Point(131, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(119, 28);
            this.label3.TabIndex = 8;
            this.label3.Text = "New Table";
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
            this.Controls.Add(this.TableName);
            this.Controls.Add(this.button1);
            this.Name = "TableForm";
            this.Text = "TableForm";
            this.ResumeLayout(false);
            this.PerformLayout();

    }

    #endregion

    private System.Windows.Forms.Button button1;
    private System.Windows.Forms.RichTextBox TableName;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.Label label2;
    private System.Windows.Forms.ComboBox TableSheet;
    private System.Windows.Forms.Label label3;
  }
}