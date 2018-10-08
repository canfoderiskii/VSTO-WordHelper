namespace WordHelper {
    partial class DocVarUserControl {
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
            this.DocVarDataGrid = new System.Windows.Forms.DataGridView();
            this.State = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.key = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.value = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DocVarContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.DocVarContextDelete = new System.Windows.Forms.ToolStripMenuItem();
            this.DocVarReloadButton = new System.Windows.Forms.Button();
            this.DocVarConfirmButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.DocVarDataGrid)).BeginInit();
            this.DocVarContextMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // DocVarDataGrid
            // 
            this.DocVarDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DocVarDataGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.State,
            this.key,
            this.value});
            this.DocVarDataGrid.ContextMenuStrip = this.DocVarContextMenu;
            this.DocVarDataGrid.Location = new System.Drawing.Point(4, 42);
            this.DocVarDataGrid.Name = "DocVarDataGrid";
            this.DocVarDataGrid.RowHeadersVisible = false;
            this.DocVarDataGrid.RowTemplate.Height = 23;
            this.DocVarDataGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.DocVarDataGrid.Size = new System.Drawing.Size(243, 395);
            this.DocVarDataGrid.TabIndex = 0;
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
            this.key.HeaderText = "变量名";
            this.key.Name = "key";
            // 
            // value
            // 
            this.value.HeaderText = "变量值";
            this.value.Name = "value";
            // 
            // DocVarContextMenu
            // 
            this.DocVarContextMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.DocVarContextDelete});
            this.DocVarContextMenu.Name = "DocVarContextMenu";
            this.DocVarContextMenu.Size = new System.Drawing.Size(125, 26);
            // 
            // DocVarContextDelete
            // 
            this.DocVarContextDelete.Name = "DocVarContextDelete";
            this.DocVarContextDelete.Size = new System.Drawing.Size(124, 22);
            this.DocVarContextDelete.Text = "删除条目";
            this.DocVarContextDelete.Click += new System.EventHandler(this.DocVarContextDelete_Click);
            // 
            // DocVarReloadButton
            // 
            this.DocVarReloadButton.Location = new System.Drawing.Point(4, 8);
            this.DocVarReloadButton.Name = "DocVarReloadButton";
            this.DocVarReloadButton.Size = new System.Drawing.Size(46, 28);
            this.DocVarReloadButton.TabIndex = 1;
            this.DocVarReloadButton.Text = "刷新";
            this.DocVarReloadButton.UseVisualStyleBackColor = true;
            this.DocVarReloadButton.Click += new System.EventHandler(this.DocVarReloadButton_Click);
            // 
            // DocVarConfirmButton
            // 
            this.DocVarConfirmButton.Location = new System.Drawing.Point(56, 8);
            this.DocVarConfirmButton.Name = "DocVarConfirmButton";
            this.DocVarConfirmButton.Size = new System.Drawing.Size(46, 28);
            this.DocVarConfirmButton.TabIndex = 1;
            this.DocVarConfirmButton.Text = "确认";
            this.DocVarConfirmButton.UseVisualStyleBackColor = true;
            this.DocVarConfirmButton.Click += new System.EventHandler(this.DocVarConfirmButton_Click);
            // 
            // DocVarUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.DocVarConfirmButton);
            this.Controls.Add(this.DocVarReloadButton);
            this.Controls.Add(this.DocVarDataGrid);
            this.Name = "DocVarUserControl";
            this.Size = new System.Drawing.Size(250, 440);
            ((System.ComponentModel.ISupportInitialize)(this.DocVarDataGrid)).EndInit();
            this.DocVarContextMenu.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView DocVarDataGrid;
        private System.Windows.Forms.Button DocVarReloadButton;
        private System.Windows.Forms.Button DocVarConfirmButton;
        private System.Windows.Forms.ContextMenuStrip DocVarContextMenu;
        private System.Windows.Forms.ToolStripMenuItem DocVarContextDelete;
        private System.Windows.Forms.DataGridViewTextBoxColumn State;
        private System.Windows.Forms.DataGridViewTextBoxColumn key;
        private System.Windows.Forms.DataGridViewTextBoxColumn value;
    }
}
