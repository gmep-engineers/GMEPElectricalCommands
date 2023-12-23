﻿namespace GMEPElectricalCommands
{
  partial class noteForm
  {
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
      if (disposing && (components != null))
      {
        components.Dispose();
      }
      base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
      this.NOTES_TEXTBOX = new System.Windows.Forms.TextBox();
      this.NOTE_PER_LINE_TEXT = new System.Windows.Forms.Label();
      this.QUICK_ADD_NOTE_TEXT = new System.Windows.Forms.Label();
      this.QUICK_ADD_COMBOBOX = new System.Windows.Forms.ComboBox();
      this.ADD_NOTE_BUTTON = new System.Windows.Forms.Button();
      this.SuspendLayout();
      // 
      // NOTES_TEXTBOX
      // 
      this.NOTES_TEXTBOX.Location = new System.Drawing.Point(12, 37);
      this.NOTES_TEXTBOX.Multiline = true;
      this.NOTES_TEXTBOX.Name = "NOTES_TEXTBOX";
      this.NOTES_TEXTBOX.Size = new System.Drawing.Size(776, 365);
      this.NOTES_TEXTBOX.TabIndex = 0;
      // 
      // NOTE_PER_LINE_TEXT
      // 
      this.NOTE_PER_LINE_TEXT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
      this.NOTE_PER_LINE_TEXT.AutoSize = true;
      this.NOTE_PER_LINE_TEXT.Location = new System.Drawing.Point(222, 21);
      this.NOTE_PER_LINE_TEXT.Name = "NOTE_PER_LINE_TEXT";
      this.NOTE_PER_LINE_TEXT.Size = new System.Drawing.Size(369, 13);
      this.NOTE_PER_LINE_TEXT.TabIndex = 1;
      this.NOTE_PER_LINE_TEXT.Text = "ENTER ONE NOTE PER LINE, PRESS ENTER TO START A NEW NOTE.";
      // 
      // QUICK_ADD_NOTE_TEXT
      // 
      this.QUICK_ADD_NOTE_TEXT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
      this.QUICK_ADD_NOTE_TEXT.AutoSize = true;
      this.QUICK_ADD_NOTE_TEXT.Location = new System.Drawing.Point(361, 414);
      this.QUICK_ADD_NOTE_TEXT.Name = "QUICK_ADD_NOTE_TEXT";
      this.QUICK_ADD_NOTE_TEXT.Size = new System.Drawing.Size(66, 13);
      this.QUICK_ADD_NOTE_TEXT.TabIndex = 2;
      this.QUICK_ADD_NOTE_TEXT.Text = "QUICK ADD";
      // 
      // QUICK_ADD_COMBOBOX
      // 
      this.QUICK_ADD_COMBOBOX.FormattingEnabled = true;
      this.QUICK_ADD_COMBOBOX.Items.AddRange(new object[] {
            "RELOCATED BREAKERS",
            "KITCHEN DEMAND FACTOR"});
      this.QUICK_ADD_COMBOBOX.Location = new System.Drawing.Point(12, 430);
      this.QUICK_ADD_COMBOBOX.Name = "QUICK_ADD_COMBOBOX";
      this.QUICK_ADD_COMBOBOX.Size = new System.Drawing.Size(666, 21);
      this.QUICK_ADD_COMBOBOX.TabIndex = 3;
      // 
      // ADD_NOTE_BUTTON
      // 
      this.ADD_NOTE_BUTTON.Location = new System.Drawing.Point(685, 429);
      this.ADD_NOTE_BUTTON.Name = "ADD_NOTE_BUTTON";
      this.ADD_NOTE_BUTTON.Size = new System.Drawing.Size(103, 23);
      this.ADD_NOTE_BUTTON.TabIndex = 4;
      this.ADD_NOTE_BUTTON.Text = "ADD NOTE";
      this.ADD_NOTE_BUTTON.UseVisualStyleBackColor = true;
      // 
      // NOTES_FORM
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(800, 464);
      this.Controls.Add(this.ADD_NOTE_BUTTON);
      this.Controls.Add(this.QUICK_ADD_COMBOBOX);
      this.Controls.Add(this.QUICK_ADD_NOTE_TEXT);
      this.Controls.Add(this.NOTE_PER_LINE_TEXT);
      this.Controls.Add(this.NOTES_TEXTBOX);
      this.Name = "NOTES_FORM";
      this.Text = "Notes";
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion

    private System.Windows.Forms.TextBox NOTES_TEXTBOX;
    private System.Windows.Forms.Label NOTE_PER_LINE_TEXT;
    private System.Windows.Forms.Label QUICK_ADD_NOTE_TEXT;
    private System.Windows.Forms.ComboBox QUICK_ADD_COMBOBOX;
    private System.Windows.Forms.Button ADD_NOTE_BUTTON;
  }
}