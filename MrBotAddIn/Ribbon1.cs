using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using MrBotAddIn.Properties;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Simple.OData.Client;
using System.Drawing;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Globalization;
using System.Threading;

namespace MrBotAddIn
{
    public partial class Ribbon1
    {
        funcionesEspeciales fe = new funcionesEspeciales();
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            connections connections = new connections();
            connections.ShowDialog();
        }
        
        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            if (fe.ReadProperty("Pagination", activeSheet.CustomProperties) != null)
            {
                Resultados formResultados = new Resultados();
                formResultados.ShowDialog();
            }
            else
            {
                MessageBox.Show("You don't have a conecction in this worksheet.");
            }
        }
        
        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            if (fe.ReadProperty("Pagination", activeSheet.CustomProperties) != null)
            {
                try
                {
                    Resultados res = new Resultados();
                    res.resetInformation();
                    Dictionary<string, int> pagination = JsonConvert.DeserializeObject<Dictionary<string, int>>(fe.ReadProperty("Pagination", activeSheet.CustomProperties));
                    pagination["whoCallsTheFunction"] = 2;
                    fe.setProperty("Pagination", activeSheet.CustomProperties, JsonConvert.SerializeObject(pagination));

                    downloadData dn = new downloadData();
                    dn.ShowDialog();
                }
                catch
                {
                    MessageBox.Show("Problem loading the data.");
                }
            }
            else
            {
                MessageBox.Show("You don't have a conecction in this worksheet.");
            }
        }

        private void button2_Click_1(object sender, RibbonControlEventArgs e)
        {
            this.RibbonUI.ActivateTab(this.Tabs[1].ControlId.ToString());
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            about ab = new about();
            ab.ShowDialog();
        }
               

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            if (fe.ReadProperty("Pagination", activeSheet.CustomProperties) != null)
            {
                activeSheet.Unprotect("123");
                connections cnt = new connections();
                List<int> numberOfInserts = JsonConvert.DeserializeObject<List<int>>(fe.ReadProperty("numberOfInserts", activeSheet.CustomProperties));
                Dictionary<string, int> pagination = JsonConvert.DeserializeObject<Dictionary<string, int>>(fe.ReadProperty("Pagination", activeSheet.CustomProperties));
                
                if ((pagination["showing"] + pagination["amountPerPage"]) < pagination["allRows"])
                {
                    pagination["whoCallsTheFunction"] = 1;
                    pagination["showing"] = pagination["showing"] + pagination["amountPerPage"];
                    if (numberOfInserts.Count == 0)
                    {
                        
                        fe.setProperty("Pagination", activeSheet.CustomProperties, JsonConvert.SerializeObject(pagination));
                        //Put the trigger in false
                        //fe.setProperty("doSomething", activeSheet.CustomProperties, JsonConvert.SerializeObject("false"));
                        downloadData x = new downloadData();
                        x.ShowDialog();
                    }
                    else
                    {
                        DialogResult dr = MessageBox.Show("You have new rows, if you show more information you will lose the registered information. Are you sure ?.", "Conflict", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information);

                        if (dr == DialogResult.Yes)
                        {
                            resetInsertFields(numberOfInserts);
                            numberOfInserts.Clear();
                            fe.setProperty("numberOfInserts", activeSheet.CustomProperties, JsonConvert.SerializeObject(numberOfInserts));
                           
                            fe.setProperty("Pagination", activeSheet.CustomProperties, JsonConvert.SerializeObject(pagination));
                            //Put the trigger in false
                            //fe.setProperty("doSomething", activeSheet.CustomProperties, JsonConvert.SerializeObject("false"));
                            downloadData x = new downloadData();
                            x.ShowDialog();
                        }
                    }
                }
                else
                {
                    MessageBox.Show("You are in the last page.");
                }
                
            }
            else
            {
                MessageBox.Show("You don't have a conecction in this worksheet.");
            }
        }

        public void resetInsertFields(List<int> numbersOfInserts)
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            List<string> columnasDeTabla = JsonConvert.DeserializeObject<List<string>>(fe.ReadProperty("columnasDeTabla", activeSheet.CustomProperties));
            foreach (int number in numbersOfInserts)
            {
                var firstCell = activeSheet.Cells[number, 1];
                var lastCell = activeSheet.Cells[number, columnasDeTabla.Count];
                var range = activeSheet.Range[firstCell, lastCell];
                range.Borders.Color = Color.FromArgb(163, 163, 163);
                range.Locked = true;
            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            if (fe.ReadProperty("Pagination", activeSheet.CustomProperties) != null)
            {
                Microsoft.Office.Interop.Excel.Range rng = (Microsoft.Office.Interop.Excel.Range)Globals.ThisAddIn.Application.ActiveCell;
                connections con = new connections();
                con.identificandoCambios(rng, "delete");
            }
            else
            {
                MessageBox.Show("You don't have a conecction in this worksheet.");
            }
            

        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            if (fe.ReadProperty("Pagination", activeSheet.CustomProperties) != null)
            {
                activeSheet.Unprotect("123");
                List<List<string>> listaDeDatosCrudosNew = JsonConvert.DeserializeObject<List<List<string>>>(fe.ReadProperty("listaDeDatosCrudosNew", activeSheet.CustomProperties));
                List<string> columnasDeTabla = JsonConvert.DeserializeObject<List<string>>(fe.ReadProperty("columnasDeTabla", activeSheet.CustomProperties));
                List<int> numberOfInserts = JsonConvert.DeserializeObject<List<int>>(fe.ReadProperty("numberOfInserts", activeSheet.CustomProperties));
                int quantity = 0;
                if (numberOfInserts.Count == 0)
                {
                    quantity = listaDeDatosCrudosNew.Count + 2;
                }
                else
                {
                    quantity = numberOfInserts.Last() + 1;
                }
                var firstCell = activeSheet.Cells[quantity, 1];
                var lastCell = activeSheet.Cells[quantity, columnasDeTabla.Count];

                var range = activeSheet.Range[firstCell, lastCell];
                range.Borders.Color = Color.FromArgb(21, 106, 75);
                range.Locked = false;

                var range2 = activeSheet.Range[firstCell, firstCell];
                range2.Select();
                numberOfInserts.Add(quantity);
                fe.setProperty("numberOfInserts", activeSheet.CustomProperties, JsonConvert.SerializeObject(numberOfInserts));

                activeSheet.Protect("123");
            }
            else
            {
                MessageBox.Show("You don't have a conecction in this worksheet.");
            }

        }

        
    }
}