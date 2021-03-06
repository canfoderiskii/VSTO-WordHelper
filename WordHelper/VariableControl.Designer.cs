﻿namespace WordHelper {
    partial class VariableControl {
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
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.VariableDataGrid = new System.Windows.Forms.DataGridView();
            this.State = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.key = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.value = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.VariableContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.VariableContextDelete = new System.Windows.Forms.ToolStripMenuItem();
            this.VariableReloadButton = new System.Windows.Forms.Button();
            this.VariableConfirmButton = new System.Windows.Forms.Button();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            ((System.ComponentModel.ISupportInitialize)(this.VariableDataGrid)).BeginInit();
            this.VariableContextMenu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.SuspendLayout();
            // 
            // VariableDataGrid
            // 
            this.VariableDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.VariableDataGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.State,
            this.key,
            this.value});
            this.VariableDataGrid.ContextMenuStrip = this.VariableContextMenu;
            this.VariableDataGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.VariableDataGrid.Location = new System.Drawing.Point(0, 0);
            this.VariableDataGrid.Name = "VariableDataGrid";
            this.VariableDataGrid.RowHeadersVisible = false;
            this.VariableDataGrid.RowTemplate.Height = 23;
            this.VariableDataGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.VariableDataGrid.Size = new System.Drawing.Size(229, 403);
            this.VariableDataGrid.TabIndex = 0;
            // 
            // State
            // 
            this.State.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.State.HeaderText = "状态";
            this.State.Name = "State";
            this.State.ReadOnly = true;
            this.State.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.State.Width = 25;
            // 
            // key
            // 
            this.key.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.key.HeaderText = "变量名";
            this.key.Name = "key";
            // 
            // value
            // 
            this.value.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.value.HeaderText = "变量值";
            this.value.Name = "value";
            // 
            // VariableContextMenu
            // 
            this.VariableContextMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.VariableContextDelete});
            this.VariableContextMenu.Name = "DocVarContextMenu";
            this.VariableContextMenu.Size = new System.Drawing.Size(125, 26);
            // 
            // VariableContextDelete
            // 
            this.VariableContextDelete.Name = "VariableContextDelete";
            this.VariableContextDelete.Size = new System.Drawing.Size(124, 22);
            this.VariableContextDelete.Text = "删除条目";
            this.VariableContextDelete.Click += new System.EventHandler(this.VariableContextDelete_Click);
            // 
            // VariableReloadButton
            // 
            this.VariableReloadButton.Location = new System.Drawing.Point(3, 3);
            this.VariableReloadButton.Name = "VariableReloadButton";
            this.VariableReloadButton.Size = new System.Drawing.Size(46, 28);
            this.VariableReloadButton.TabIndex = 1;
            this.VariableReloadButton.Text = "刷新";
            this.VariableReloadButton.UseVisualStyleBackColor = true;
            this.VariableReloadButton.Click += new System.EventHandler(this.VariableReloadButton_Click);
            // 
            // VariableConfirmButton
            // 
            this.VariableConfirmButton.Location = new System.Drawing.Point(55, 3);
            this.VariableConfirmButton.Name = "VariableConfirmButton";
            this.VariableConfirmButton.Size = new System.Drawing.Size(46, 28);
            this.VariableConfirmButton.TabIndex = 1;
            this.VariableConfirmButton.Text = "确认";
            this.VariableConfirmButton.UseVisualStyleBackColor = true;
            this.VariableConfirmButton.Click += new System.EventHandler(this.VariableConfirmButton_Click);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.IsSplitterFixed = true;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.VariableReloadButton);
            this.splitContainer1.Panel1.Controls.Add(this.VariableConfirmButton);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.VariableDataGrid);
            this.splitContainer1.Size = new System.Drawing.Size(229, 440);
            this.splitContainer1.SplitterDistance = 33;
            this.splitContainer1.TabIndex = 2;
            // 
            // VariableControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.splitContainer1);
            this.Name = "VariableControl";
            this.Size = new System.Drawing.Size(229, 440);
            ((System.ComponentModel.ISupportInitialize)(this.VariableDataGrid)).EndInit();
            this.VariableContextMenu.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView VariableDataGrid;
        private System.Windows.Forms.Button VariableReloadButton;
        private System.Windows.Forms.Button VariableConfirmButton;
        private System.Windows.Forms.ContextMenuStrip VariableContextMenu;
        private System.Windows.Forms.ToolStripMenuItem VariableContextDelete;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.DataGridViewTextBoxColumn State;
        private System.Windows.Forms.DataGridViewTextBoxColumn key;
        private System.Windows.Forms.DataGridViewTextBoxColumn value;
    }
}
