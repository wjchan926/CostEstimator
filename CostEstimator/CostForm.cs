﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Inventor;

namespace InvAddIn
{
    public partial class CostForm : Form
    {
        Inventor.Application invApp;

        CostForm()
        {

        }

        public CostForm(Inventor.Application currentApp)
        {
            invApp = currentApp;
            InitializeComponent();
        }


        private void CostForm_Load(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void updateCostBtn_Click(object sender, EventArgs e)
        {
            // Do nothing if no file open
            try
            {
                Costable currentFile = new Costable(invApp, invApp.ActiveDocument);
                currentFile.CostOf(currentFile.inventorDoc);

                partNumLabel.Text = currentFile.inventorDoc.DisplayName;

                costLabel.Text = string.Format("{0:C}", Convert.ToDecimal(currentFile.cost));

                currentFile.UpdateMaterial(currentFile.inventorDoc);
                matGridView.DataSource = currentFile.matBreakdown;
                matGridView.Update();
                matGridView.Refresh();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
    //rawCostLabel.Text = string.Format("{0:C}", Convert.ToDecimal(currentFile.cost));
    //purchasedCostLabel.Text = string.Format("{0:C}", Convert.ToDecimal(currentFile.cost));
}
    }
}
