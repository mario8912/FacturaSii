﻿using System.Windows.Forms;

namespace Presentacion
{
    partial class Form1
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
            this.btnCrearXml = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btnSeleccionArchivo = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnCrearXml
            // 
            this.btnCrearXml.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCrearXml.Enabled = false;
            this.btnCrearXml.Location = new System.Drawing.Point(500, 127);
            this.btnCrearXml.Name = "btnCrearXml";
            this.btnCrearXml.Size = new System.Drawing.Size(91, 22);
            this.btnCrearXml.TabIndex = 0;
            this.btnCrearXml.Text = "Crear XML";
            this.btnCrearXml.UseVisualStyleBackColor = true;
            this.btnCrearXml.Click += new System.EventHandler(this.btnCrearXml_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(141, 35);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(319, 20);
            this.textBox1.TabIndex = 1;
            // 
            // btnSeleccionArchivo
            // 
            this.btnSeleccionArchivo.Location = new System.Drawing.Point(6, 35);
            this.btnSeleccionArchivo.Name = "btnSeleccionArchivo";
            this.btnSeleccionArchivo.Size = new System.Drawing.Size(129, 20);
            this.btnSeleccionArchivo.TabIndex = 3;
            this.btnSeleccionArchivo.Text = "Seleccionar Archivo";
            this.btnSeleccionArchivo.UseVisualStyleBackColor = true;
            this.btnSeleccionArchivo.Click += new System.EventHandler(this.botonSelecionArchivo_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.btnSeleccionArchivo);
            this.groupBox1.Controls.Add(this.btnCrearXml);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(597, 155);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Factura Sii";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(622, 179);
            this.Controls.Add(this.groupBox1);
            this.Name = "Form1";
            this.Text = "Sii";
            this.Activated += new System.EventHandler(this.Form1_Activated);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnCrearXml;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btnSeleccionArchivo;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}