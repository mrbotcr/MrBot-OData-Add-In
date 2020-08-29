using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Excel;
using Simple.OData.Client;

namespace MrBotAddIn
{
    public partial class Resultados : Form
    {
        funcionesEspeciales fe = new funcionesEspeciales();
        conexionesOData conexionSeleccionada;
        ODataClient client;
        connectHTTPS cnttps = new connectHTTPS();
        //private delegate void MessageFromProcessDelegate(string message, Color color);
        //private delegate void StopDelegate();
        public Resultados()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Resultados_Load(object sender, EventArgs e)
        {
            MostrarDetallesAEditar();
        }

        public void ChangeText(string newText, Color color)
        {
            richTextBox1.SelectionColor = color;
            richTextBox1.AppendText(newText);
        }

        public void MostrarDetallesAEditar()
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            List<changeInformation> listOfChanges = JsonConvert.DeserializeObject<List<changeInformation>>(fe.ReadProperty("listOfChanges", activeSheet.CustomProperties));
            List<int> numberOfInserts = JsonConvert.DeserializeObject<List<int>>(fe.ReadProperty("numberOfInserts", activeSheet.CustomProperties));
            ChangeText("Summary of processes to be executed ...\n", Color.Black);
            var updates = listOfChanges.Where(x => x.ch_action == "update");
            var deletes = listOfChanges.Where(x => x.ch_action == "delete");
            var groupsOfUpdates = updates.GroupBy(x => x.ch_idRow);
            ChangeText("Number of lines to EDIT:"+ groupsOfUpdates.Count().ToString()+ "\n", Color.Orange);
            ChangeText("Number of lines to DELETE:"+ deletes.Count().ToString()+ "\n", Color.Orange);
            ChangeText("Number of lines to INSERT:" + numberOfInserts.Count.ToString() + "\n", Color.Orange);
            ChangeText("Click 'Start' to start the process ...\n", Color.Black);
            progressBar1.Maximum = groupsOfUpdates.Count() + deletes.Count() + numberOfInserts.Count;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //try {
                ChangeText("Starting the process ...\n", Color.Black);
                button2.Enabled = false;
                button1.Enabled = false;
                processChanges();
                                
            /*}
            catch
            {
                ChangeText("Connection not established ...\n", Color.Red);
                MessageBox.Show("Connection not established.");
            }
            */
        }

        public void processChanges()
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            List<string> columnasDeTabla = JsonConvert.DeserializeObject<List<string>>(fe.ReadProperty("columnasDeTabla", activeSheet.CustomProperties));
            conexionSeleccionada = JsonConvert.DeserializeObject<conexionesOData>(fe.ReadProperty("conexionSeleccionada", activeSheet.CustomProperties));
            client = cnttps.conectWithCredentials(conexionSeleccionada);
            List<List<string>> listaDeDatosCrudosNew = JsonConvert.DeserializeObject<List<List<string>>>(fe.ReadProperty("listaDeDatosCrudosNew", activeSheet.CustomProperties));
            List<changeInformation> listOfChanges = JsonConvert.DeserializeObject<List<changeInformation>>(fe.ReadProperty("listOfChanges", activeSheet.CustomProperties));
            List<int> numberOfInserts = JsonConvert.DeserializeObject<List<int>>(fe.ReadProperty("numberOfInserts", activeSheet.CustomProperties));
            conexionTabla conexiontabla = JsonConvert.DeserializeObject<conexionTabla>(fe.ReadProperty("conexionTabla", activeSheet.CustomProperties));

            var updates = listOfChanges.Where(x => x.ch_action == "update");
            var deletes = listOfChanges.Where(x => x.ch_action == "delete");

            var groupsOfUpdates = updates.GroupBy(x => x.ch_idRow);

            if(groupsOfUpdates.Count() > 0)
                ChangeText("Updating information...\n", Color.Black);

            foreach (var group in groupsOfUpdates)
            {
                IDictionary<string, object> setInfo = new Dictionary<string, object>();
                
                int idRowExcel = 0;
                List<llavesPrimariasClass> llaves = new List<llavesPrimariasClass>();
                foreach (var column in group)
                {
                    idRowExcel = column.ch_idRowExcel;
                    llaves = column.ch_llaves_primarias;
                    setInfo.Add(column.ch_column, column.ch_newValue);
                }

                aplicarUpdate(conexiontabla.Tabla,setInfo, idRowExcel, convertKeysToFilter(llaves), llaves.Count, client);
                
            }

            if (groupsOfUpdates.Count() > 0)
                ChangeText("Finished ...\n", Color.Black);

            if (deletes.Count() > 0)
                ChangeText("Deleting information ...\n", Color.Black);

            foreach (var delete in deletes)
            {
                aplicarDelete(conexiontabla.Tabla, delete.ch_idRowExcel, convertKeysToFilter(delete.ch_llaves_primarias), delete.ch_llaves_primarias.Count, client);
            }

            if (deletes.Count() > 0)
                ChangeText("Finished ...\n", Color.Black);

            if (numberOfInserts.Count() > 0)
                ChangeText("Inserting information ...\n", Color.Black);

            foreach (int row in numberOfInserts)
            {
                getInformationFromExcel(columnasDeTabla, row, conexiontabla.Tabla);
            }

            if (numberOfInserts.Count() > 0)
                ChangeText("Finished ...\n", Color.Black);
            
            ChangeText("Procedure completed successfully ...\nClick 'Close' to return to the Excel...\n", Color.Green);
            resetInformation();
            button1.Enabled = true;

        }

        public void resetInformation()
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            List<changeInformation> listOfChanges = new List<changeInformation>();
            List<changeInformation> listOfChangesToRecover = new List<changeInformation>();
            List<int> numberOfInserts = new List<int>();
            fe.setProperty("listOfChanges", activeSheet.CustomProperties, JsonConvert.SerializeObject(listOfChanges));
            fe.setProperty("listOfChangesToRecover", activeSheet.CustomProperties, JsonConvert.SerializeObject(listOfChangesToRecover));
            fe.setProperty("numberOfInserts", activeSheet.CustomProperties, JsonConvert.SerializeObject(numberOfInserts));
        }

        public void aplicarUpdate(string tabla, IDictionary<string, object> set, int idRow, object llave, int canLlaves, ODataClient client)
        {
            Task.Run(async () =>
            {
                try
                {
                    //var ship2 = new Task(() => { });
                    IDictionary<string, object> ship;
                    if (canLlaves == 1)
                    {
                        await client.For(tabla).Key(llave).Set(set).UpdateEntryAsync();
                    }
                    else
                    {
                        await client.For(tabla).Filter(llave.ToString()).Set(set).UpdateEntryAsync();
                    }
                    ChangeText("Line: " + idRow.ToString() + " - UPDATE - Correct.\n", Color.Green);
                }
                catch (WebRequestException ex)
                {
                    try
                    {
                        ChangeText("Line: " + idRow.ToString() + " - UPDATE - Error.\n", Color.Red);
                        dynamic obj = JsonConvert.DeserializeObject(ex.Response);
                        var messageFromServer = obj.error.message.value.ToString();
                        ChangeText("Message: " + messageFromServer + "\n", Color.Red);
                    }
                    catch
                    {
                        ChangeText("Line: " + idRow.ToString() + " - UPDATE - Error.\n", Color.Red);
                    }
                }
                catch (Exception e)
                {
                    ChangeText("Line: " + idRow.ToString() + " - UPDATE - Error.\n", Color.Red);
                }
                progressBar1.Value = progressBar1.Value + 1;
            }).Wait();
        }

        public void aplicarDelete(string tabla, int idRow, object llave, int canLlaves, ODataClient client)
        {
            Task.Run(async () =>
            {
                try
                {
                    if (canLlaves == 1)
                    {
                        var ship = await client.For(tabla).Key(llave).DeleteEntriesAsync();
                    }
                    else
                    {
                        var ship = await client.For(tabla).Filter(llave.ToString()).DeleteEntriesAsync();
                    }
                    ChangeText("Line: " + idRow.ToString() + " - DELETE - Correct.\n", Color.Green);
                }
                catch (WebRequestException ex)
                {

                    try
                    {
                        ChangeText("Line: " + idRow.ToString() + " - DELETE - Error.\n", Color.Red);
                        dynamic obj = JsonConvert.DeserializeObject(ex.Response);
                        var messageFromServer = obj.error.message.value.ToString();
                        ChangeText("Message: " + messageFromServer + "\n", Color.Red);
                    }
                    catch
                    {
                        ChangeText("Line: " + idRow.ToString() + " - DELETE - Error.\n", Color.Red);
                    }
                }
                catch (Exception e)
                {
                    ChangeText("Line: " + idRow.ToString() + " - DELETE - Error.\n", Color.Red);
                }
                progressBar1.Value = progressBar1.Value + 1;
            }).Wait();
        }

        public void getInformationFromExcel(List<string> columnas, int row, string tabla)
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            try
            {
                List<string> datos_act = new List<string>();
                List<int> listaDeFechas = JsonConvert.DeserializeObject<List<int>>(fe.ReadProperty("ListaDeDatetime", activeSheet.CustomProperties));
                Dictionary<string, object> dataForInsert = new Dictionary<string, object>();

                for (int col = 1; col <= columnas.Count; col++)
                {
                    var data_ = ((Microsoft.Office.Interop.Excel.Range)activeSheet.Cells[row, col]).Value2;

                    if (data_ is double)
                    {
                        //MessageBox.Show(data_.ToString());
                        if (listaDeFechas.Any(x => x == col))
                        {
                            DateTime dt = DateTime.FromOADate((double)data_);
                            data_ = dt.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        else
                        {
                            var data_2 = ((Microsoft.Office.Interop.Excel.Range)activeSheet.Cells[row, col]).Text;
                            try
                            {
                                DateTime dateTime = DateTime.Parse(data_2);
                                data_ = dateTime.ToString("yyyy-MM-dd HH:mm:ss");
                            }
                            catch
                            {
                                data_ = ((double)((Microsoft.Office.Interop.Excel.Range)activeSheet.Cells[row, col]).Value2).ToString();
                            }

                        }
                    }
                    dataForInsert.Add(columnas.ElementAt(col-1), data_);
                }
                aplicarInsert(row,dataForInsert, tabla);
            }
            catch
            {
                ChangeText("Line " + row.ToString() + " - INSERT - Error.\n", Color.Red);
                progressBar1.Value = progressBar1.Value + 1;
            }
        }

        public async void aplicarInsert(int row, Dictionary<string, object> information, string tabla)
        {
            Task.Run(async () =>
            {
                try
                {
                    var ship = await client.For(tabla).Set(information).InsertEntryAsync();
                    ChangeText("Line: " + row.ToString() + " - INSERT - Correct.\n", Color.Green);
                }
                catch (WebRequestException ex)
                {
                    try
                    {
                        ChangeText("Line " + row.ToString() + " - INSERT - Error.\n", Color.Red);
                        dynamic obj = JsonConvert.DeserializeObject(ex.Response);
                        var messageFromServer = obj.error.message.value.ToString();
                        ChangeText("Message: " + messageFromServer + "\n", Color.Red);
                    }
                    catch
                    {
                        ChangeText("Line " + row.ToString() + " - INSERT - Error.\n", Color.Red);
                    }
                }
                catch (Exception ex)
                {
                    ChangeText("Line " + row.ToString() + " - INSERT - Error.\n", Color.Red);
                }

                progressBar1.Value = progressBar1.Value + 1;
            }).Wait();
        }


        public object convertKeysToFilter(List<llavesPrimariasClass> llavesPrimarias)
        {
            try
            {
                if (llavesPrimarias.Count == 1)
                {
                    return llavesPrimarias[0].llp_object;
                }
                else
                {
                    StringBuilder sb = new StringBuilder();
                    int cantLlaves = llavesPrimarias.Count - 1;
                    int contador = 1;
                    foreach (llavesPrimariasClass llave in llavesPrimarias)
                    {
                        sb.Append(String.Format("({0} eq {1})", llave.llp_name, llave.llp_object));
                        if (contador <= cantLlaves)
                        {
                            sb.Append(" and ");
                        }
                        contador = contador + 1;
                    }

                    return sb.ToString();
                }
            }
            catch
            {
                MessageBox.Show("KEY: convertKeysToFilter.");
                return null;
            }
        }

        private void Resultados_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (button2.Enabled == false)
            {
                this.Hide();
                Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
                Dictionary<string, int> pagination = JsonConvert.DeserializeObject<Dictionary<string, int>>(fe.ReadProperty("Pagination", activeSheet.CustomProperties));
                pagination["whoCallsTheFunction"] = 2;
                fe.setProperty("Pagination", activeSheet.CustomProperties, JsonConvert.SerializeObject(pagination));

                downloadData dn = new downloadData();
                dn.ShowDialog();
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}

class ErrorOData
{
    public Error error { get; set; }
}
