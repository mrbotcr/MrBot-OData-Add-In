using Microsoft.Data.Edm;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Simple.OData.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace MrBotAddIn
{
    public partial class connections : Form
    {
        System.Globalization.CultureInfo oldCI = new System.Globalization.CultureInfo("en-US");
        connectHTTPS cnttps = new connectHTTPS();
        funcionesEspeciales fe = new funcionesEspeciales();
        ODataClient client;
        conexionesOData conexionSeleccionada;
        Ribbon1 evente = new Ribbon1();
        Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
        public connections()
        {
            InitializeComponent();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            /*
                DELETE CONNECTION

                We obtain the list of connections of the properties of the project and convert them into a list.
                The connection with the name to be deleted is searched and removed from the list.
                The list is updated in the properties and changes are saved.
            */
            DialogResult dr = MessageBox.Show("Are you sure ?.", "Delete connection", MessageBoxButtons.YesNoCancel,MessageBoxIcon.Information);

            if (dr == DialogResult.Yes)
            {
                List<conexionesOData> lista = JsonConvert.DeserializeObject<List<conexionesOData>>(MrBotAddIn.Properties.Settings.Default.jsonDeConexiones);
                if (lista != null)
                {
                    foreach (conexionesOData connect in lista)
                    {
                        if (comboBox1.Text == connect.Name)
                        {
                            lista.Remove(connect);
                            break;
                        }
                    }
                }
                MrBotAddIn.Properties.Settings.Default.jsonDeConexiones = JsonConvert.SerializeObject(lista);
                MrBotAddIn.Properties.Settings.Default.Save();
                cargarConexiones();
            }
        }

        private void connections_Load(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;
            cargarConexiones();            
        }

        public void cargarConexiones()
        {
            /*
                We obtain the list of connections of the properties of the project, 
                convert them into a list and load them in the comboboxes.
            */
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            //MrBotAddIn.Properties.Settings.Default.Reset();
            List <conexionesOData> lista = JsonConvert.DeserializeObject<List<conexionesOData>>(MrBotAddIn.Properties.Settings.Default.jsonDeConexiones);
            if (lista != null)
            {
                foreach (conexionesOData connect in lista)
                {
                    comboBox1.Items.Add(connect.Name);
                }
            }
            comboBox1.Items.Add("New");
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*
                Every time the connection combobox changes an action is made.
                If the selected option is "New" then the interface to create a new connection is opened.
                If the option is one of the connections already created, it proceeds to obtain the tables it has. And they are loaded in the Tables combobox.
            */
            try
            {

                if (comboBox1.Text == "New")
                {
                    nuevaConexion nuevaConexion = new nuevaConexion();
                    nuevaConexion.Owner = this;
                    nuevaConexion.ShowDialog();
                } else
                {
                    comboBox2.Text = "";
                    List<conexionesOData> lista = JsonConvert.DeserializeObject<List<conexionesOData>>(MrBotAddIn.Properties.Settings.Default.jsonDeConexiones);
                    conexionSeleccionada = lista.Find(x => x.Name == comboBox1.Text);
                    client = cnttps.conectWithCredentials(conexionSeleccionada);
                    showTablesInComboBox(client);
                }
            }catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        
        public async void showTablesInComboBox(ODataClient client)
        {
            try {
                comboBox2.Items.Clear();
                IEdmModel metadata = await client.GetMetadataAsync<IEdmModel>();

                var entityTypes = metadata.SchemaElements.OfType<IEdmEntityType>().ToArray();
                foreach (var type in entityTypes)
                {
                    comboBox2.Items.Add(type.Name);
                }
            }catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //It is checked that a valid table is selected in the table combobox
            if (comboBox2.Text != "")
            {
                try {
                    //A worksheet is created to show the data and assign the necessary variables
                    activeSheet = (Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
                    //Set the name of the active sheet
                    if(comboBox2.Text.Length >27)
                    {
                        activeSheet.Name = comboBox2.Text.Substring(0, 27);
                    }
                    else
                    {
                        activeSheet.Name = comboBox2.Text;
                    }
                    
                   
                    //Conexion OData Selected
                    activeSheet.CustomProperties.Add("conexionSeleccionada", JsonConvert.SerializeObject(conexionSeleccionada));
                    //Paginacion
                    Dictionary<string, int> pagination = new Dictionary<string, int>();
                    pagination["whoCallsTheFunction"] = 0;
                    pagination["amountPerPage"] = Convert.ToInt32(numericUpDown1.Value);
                    pagination["showing"] = 0;
                    pagination["allRows"] = 0;
                    activeSheet.CustomProperties.Add("Pagination", JsonConvert.SerializeObject(pagination));
                    //Know the table we are working
                    conexionTabla conexiontabla = new conexionTabla();
                    conexiontabla.Url = conexionSeleccionada.Url;
                    conexiontabla.Tabla = comboBox2.Text;
                    activeSheet.CustomProperties.Add("conexionTabla", JsonConvert.SerializeObject(conexiontabla));
                    //know date type fields
                    List<int> listOfIdOfDatetime = new List<int>();
                    activeSheet.CustomProperties.Add("ListaDeDatetime", JsonConvert.SerializeObject(listOfIdOfDatetime));
                    //We create the list of list for the raw data
                    List<List<string>> listaDeDatosCrudosNew = new List<List<string>>();
                    activeSheet.CustomProperties.Add("listaDeDatosCrudosNew", JsonConvert.SerializeObject(listaDeDatosCrudosNew));
                    //We create a list for register change
                    List<changeInformation> listOfChanges = new List<changeInformation>();
                    activeSheet.CustomProperties.Add("listOfChanges", JsonConvert.SerializeObject(listOfChanges));
                    List<changeInformation> listOfChangesToRecover = new List<changeInformation>();
                    activeSheet.CustomProperties.Add("listOfChangesToRecover", JsonConvert.SerializeObject(listOfChangesToRecover));
                    activeSheet.CustomProperties.Add("doSomething", JsonConvert.SerializeObject("true"));
                    activeSheet.CustomProperties.Add("crearEvento", JsonConvert.SerializeObject("si"));
                    List<int> numberOfInserts = new List<int>();
                    activeSheet.CustomProperties.Add("numberOfInserts", JsonConvert.SerializeObject(numberOfInserts));
                    iniciarDatos();
                    this.Close();
                }catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("You must select a table to show the data.");
            }

        }

        public void iniciarDatos()
        {
            try {
                //You always work on the active sheet
                
                //For put the information in red when change the field 
                //activeSheet.Change += WorksheetChangeEventHandler;
                DocEvents_ChangeEventHandler EventDel_CellsChange = new DocEvents_ChangeEventHandler(WorksheetChangeEventHandler);
                activeSheet.Change += EventDel_CellsChange;
                //The established connection of the properties of the sheet is obtained
                conexionesOData conexionSeleccionada = JsonConvert.DeserializeObject<conexionesOData>(fe.ReadProperty("conexionSeleccionada", activeSheet.CustomProperties));
                //The OData connection is established
                client = cnttps.conectWithCredentials(conexionSeleccionada);
                //Download
                downloadData down = new downloadData();
                down.ShowDialog();
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        public void DarFormatoALlavesPrimarias(List<int>llavesPrimarias, Worksheet activeSheet, int cantidadDeDatos)
        {
            foreach (int llavePri in llavesPrimarias)
            {
                var firstCell2 = activeSheet.Cells[2, llavePri + 1];
                var lastCell2 = activeSheet.Cells[cantidadDeDatos + 1, llavePri + 1];

                var range2 = activeSheet.Range[firstCell2, lastCell2];
                range2.Font.Color = Color.FromArgb(218, 165, 32);
                range2.Font.Bold = true;
                range2.Locked = true;
            }
        }

        public void WorksheetChangeEventHandler(Range Target)
        {
            try
            {
                //Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
                string trigger = JsonConvert.DeserializeObject<string>(fe.ReadProperty("doSomething", activeSheet.CustomProperties));
                if (trigger =="true")
                {
                    identificandoCambios(Target, "update");
                }
                else
                {
                    trigger = "true";
                    fe.setProperty("doSomething", activeSheet.CustomProperties, JsonConvert.SerializeObject(trigger));
                }

            }
            catch (Exception e)
            {
            }
        }

        public void identificandoCambios(Range Target, string actionFunc)
        {
            //Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            activeSheet.Unprotect("123");
            List<List<string>> listaDeDatosCrudosNew = JsonConvert.DeserializeObject<List<List<string>>>(fe.ReadProperty("listaDeDatosCrudosNew", activeSheet.CustomProperties));
            List<string> columnasDeTabla = JsonConvert.DeserializeObject<List<string>>(fe.ReadProperty("columnasDeTabla", activeSheet.CustomProperties));
            int row2 = Target.Row;
            if (listaDeDatosCrudosNew.Count + 1 >= row2)
            {
                List<int> llavesPrimarias = JsonConvert.DeserializeObject<List<int>>(fe.ReadProperty("idLlavePrimaria", activeSheet.CustomProperties));
                List<int> listaDeFechas = JsonConvert.DeserializeObject<List<int>>(fe.ReadProperty("ListaDeDatetime", activeSheet.CustomProperties));
                List<changeInformation> listOfChanges = JsonConvert.DeserializeObject<List<changeInformation>>(fe.ReadProperty("listOfChanges", activeSheet.CustomProperties));
                List<changeInformation> listOfChangesToRecover = JsonConvert.DeserializeObject<List<changeInformation>>(fe.ReadProperty("listOfChangesToRecover", activeSheet.CustomProperties));
                List<llavesPrimariasClass> llavesPrimariasClass = new List<llavesPrimariasClass>();

                try
                {
                    int row = Target.Row - 2;
                    int col = Target.Column - 1;
                    var datosNew = Target.Value2;
                    bool saber = datosNew is Array;
                    if (!saber)
                    {
                        List<string> infoBuscando = new List<string>();
                        foreach (int idLlavePrimaria in llavesPrimarias)
                        {
                            int columnaid = idLlavePrimaria + 1;
                            var idBuscando = ((Microsoft.Office.Interop.Excel.Range)activeSheet.Cells[row2, columnaid]).Value2;
                            if (idBuscando is double)
                            {
                                if (listaDeFechas.Any(x => x == columnaid))
                                {
                                    DateTime dt = DateTime.FromOADate((double)idBuscando);
                                    idBuscando = dt.ToString("yyyy-MM-dd HH:mm:ss");
                                }

                            }
                            infoBuscando.Add(idBuscando.ToString());
                            var llp = new llavesPrimariasClass();
                            llp.llp_object = idBuscando;
                            llp.llp_name = columnasDeTabla.ElementAt(idLlavePrimaria);
                            llavesPrimariasClass.Add(llp);
                        }
                        List<List<string>> sacandoSoloLasColumnasLlave = listaDeDatosCrudosNew.Select(x => llavesPrimarias.Select(index => x[index]).ToList()).ToList();
                        int datosCrudoIndex = sacandoSoloLasColumnasLlave.Select((data, index) => new { Data = string.Join("#", data), index = index }).Where(select => select.Data == string.Join("#", infoBuscando)).Select(a => a.index).First();
                        var datosCrudo = listaDeDatosCrudosNew.ElementAt(datosCrudoIndex);
                        changeInformation new_change = new changeInformation();
                        if (actionFunc == "update")
                        {
                            var searchChangeDel = listOfChanges.FirstOrDefault(x => x.ch_idRow == datosCrudoIndex && x.ch_columnId == col);
                            if (searchChangeDel != null)
                            {
                                listOfChanges.Remove(searchChangeDel);
                            }
                        }

                        List<changeInformation> searchChange = new List<changeInformation>();
                        if (actionFunc == "delete")
                        {
                            searchChange = listOfChanges.Where(x => x.ch_idRow == datosCrudoIndex).ToList();
                            if (searchChange.Count != 0)
                            {
                                var flag = "delete";
                                foreach (changeInformation change in searchChange)
                                {
                                    listOfChanges.Remove(change);
                                    listOfChangesToRecover.Add(change);
                                    if (change.ch_column != null)
                                    {
                                        flag = "update";
                                    }
                                }
                                if (flag == "update")
                                {
                                    searchChange = listOfChanges.Where(x => x.ch_idRow == datosCrudoIndex).ToList();
                                }

                            }
                        }

                        new_change.ch_action = actionFunc;
                        new_change.ch_idRow = datosCrudoIndex;
                        new_change.ch_idRowExcel = Target.Row;
                        new_change.ch_llaves_primarias = llavesPrimariasClass;

                        if (actionFunc == "update")
                        {
                            if (datosCrudo.ElementAt(col) != Convert.ToString(datosNew))
                            {
                                Target.Cells.Font.Color = Color.FromArgb(255, 0, 0);
                                new_change.ch_column = columnasDeTabla[col];
                                new_change.ch_columnId = col;
                                new_change.ch_newValue = reviewFormat(datosNew, Target.Row, Target.Column);
                                listOfChanges.Add(new_change);
                            }
                            else
                            {
                                Target.Cells.Font.Color = Color.FromArgb(0, 0, 0);
                            }
                        }
                        else
                        {
                            var firstCell = activeSheet.Cells[Target.Row, 1];
                            var lastCell = activeSheet.Cells[Target.Row, columnasDeTabla.Count];
                            var range = activeSheet.Range[firstCell, lastCell];
                            if (searchChange.Count == 0)
                            {
                                activeSheet.Range[firstCell, lastCell].Font.Color = Color.FromArgb(255, 0, 0);
                                listOfChanges.Add(new_change);
                                range.Locked = true;
                            }
                            else
                            {
                                activeSheet.Range[firstCell, lastCell].Font.Color = Color.FromArgb(0, 0, 0);
                                range.Locked = false;
                                DarFormatoALlavesPrimarias(llavesPrimarias, activeSheet, listaDeDatosCrudosNew.Count);

                                //Recuperar datos viejos si tuviera
                                searchChange = listOfChangesToRecover.Where(x => x.ch_idRow == datosCrudoIndex).ToList();
                                foreach (changeInformation change in searchChange)
                                {
                                    listOfChangesToRecover.Remove(change);
                                    if (change.ch_column != null)
                                    {
                                        var firstCell2 = activeSheet.Cells[Target.Row, change.ch_columnId + 1];
                                        var lastCell2 = activeSheet.Cells[Target.Row, change.ch_columnId + 1];
                                        var range2 = activeSheet.Range[firstCell2, lastCell2];
                                        range2.value2 = datosCrudo.ElementAt(change.ch_columnId);
                                        listOfChangesToRecover.Remove(change);
                                    }
                                }
                            }
                        }

                        fe.setProperty("listOfChanges", activeSheet.CustomProperties, JsonConvert.SerializeObject(listOfChanges));
                        fe.setProperty("listOfChangesToRecover", activeSheet.CustomProperties, JsonConvert.SerializeObject(listOfChangesToRecover));
                    }
                }

                catch
                {
                }
            }
            else
            {
                if (actionFunc == "delete")
                {
                    
                    fe.setProperty("doSomething", activeSheet.CustomProperties, JsonConvert.SerializeObject("False"));
                    List<int> numberOfInserts = JsonConvert.DeserializeObject<List<int>>(fe.ReadProperty("numberOfInserts", activeSheet.CustomProperties));
                    numberOfInserts.Remove(row2);
                    var firstCell2 = activeSheet.Cells[Target.Row,1];
                    var lastCell2 = activeSheet.Cells[Target.Row, columnasDeTabla.Count];
                    var range2 = activeSheet.Range[firstCell2, lastCell2];
                    range2.Borders.Color = Color.FromArgb(163, 163, 163);
                    range2.value2 = "";
                    range2.Locked = true;

                    fe.setProperty("numberOfInserts", activeSheet.CustomProperties, JsonConvert.SerializeObject(numberOfInserts));
                }
            }
                activeSheet.Protect("123");
        }

        public object reviewFormat(object data_, int row, int col)
        {
            //Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            List<int> listaDeFechas = JsonConvert.DeserializeObject<List<int>>(fe.ReadProperty("ListaDeDatetime", activeSheet.CustomProperties));

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

            return data_;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void connections_EnabledChanged(object sender, EventArgs e)
        {
            /*ESTE EVENTO SE DISPARA CUANDO SE CIERRA LA VENTANA DE LA CREACIÓN DE UNA CONEXION*/
            cargarConexiones();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
