using Microsoft.Data.Edm;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Simple.OData.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MrBotAddIn
{
    public partial class downloadData : Form
    {
        funcionesEspeciales fe = new funcionesEspeciales();
        connectHTTPS cnttps = new connectHTTPS();
        connections cnt = new connections();
        System.Threading.Thread currentThread;
        ODataFeedAnnotations annotations = new ODataFeedAnnotations();
        public downloadData()
        {
            InitializeComponent();
        }

        public void RefreshProgress(int id)
        {
            if (this == null) return;
            if (id == 1)
                Invoke(new System.Action(() => progressBar1.Value = progressBar1.Value + 1));
            else
                Invoke(new System.Action(() => progressBar2.Value = progressBar2.Value + 1));
        }

        private void Start()
        {
            if (currentThread == null)
            {
                currentThread = new System.Threading.Thread(DoProcess);
                currentThread.Start();
            }
        }
        private void Stop()
        {
            if (currentThread != null)
            {
                currentThread.Abort();
                currentThread = null;
            }
            
        }

        private async void DoProcess()
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            Dictionary<string, int> pagination = JsonConvert.DeserializeObject<Dictionary<string, int>>(fe.ReadProperty("Pagination", activeSheet.CustomProperties));
            //THE INITIAL CONNECTION
            if (pagination["whoCallsTheFunction"] == 0)
            {
                fe.setProperty("doSomething", activeSheet.CustomProperties, JsonConvert.SerializeObject("false"));
                obtenerColumnasDeTabla();
                await traerDatosDeTabla();

            }
            else
            {
                //SHOW MORE ACTION BUTTON
                if (pagination["whoCallsTheFunction"] == 1)
                {
                    if (pagination["allRows"] >= pagination["showing"] + pagination["amountPerPage"])
                        progressBar1.Maximum = pagination["amountPerPage"];
                    else
                        progressBar1.Maximum = pagination["allRows"] - pagination["showing"];

                    label3.ForeColor = Color.Green;
                    activeSheet.Unprotect("123");
                    fe.setProperty("doSomething", activeSheet.CustomProperties, JsonConvert.SerializeObject("false"));
                    await traerDatosDeTabla();

                }
                else
                {
                    //REFRESH
                    if (pagination["whoCallsTheFunction"] == 2)
                    {
                        activeSheet.Unprotect("123");
                        fe.setProperty("doSomething", activeSheet.CustomProperties, JsonConvert.SerializeObject("false"));
                        activeSheet.Cells.Delete();
                        conexionTabla conexiontabla = JsonConvert.DeserializeObject<conexionTabla>(fe.ReadProperty("conexionTabla", activeSheet.CustomProperties));
                        conexionesOData conexionSeleccionada = JsonConvert.DeserializeObject<conexionesOData>(fe.ReadProperty("conexionSeleccionada", activeSheet.CustomProperties));
                        ODataClient client = cnttps.conectWithCredentials(conexionSeleccionada);
                        var informacion = await client.For(conexiontabla.Tabla).Top(1).FindEntriesAsync(annotations);
                        int newRows = Convert.ToInt32(annotations.Count);
                        List<int> listOfIdOfDatetime = new List<int>();
                        List<List<string>> listaDeDatosCrudosNew = new List<List<string>>();
                        fe.setProperty("listaDeDatosCrudosNew", activeSheet.CustomProperties, JsonConvert.SerializeObject(listaDeDatosCrudosNew));
                        fe.setProperty("ListaDeDatetime", activeSheet.CustomProperties, JsonConvert.SerializeObject(listOfIdOfDatetime));

                        int ntimes = 0;
                        int estabaMostrando = pagination["showing"] + pagination["amountPerPage"];
                        if (estabaMostrando <= newRows)
                        {
                            ntimes = estabaMostrando / pagination["amountPerPage"];
                            progressBar1.Maximum = estabaMostrando;
                        }
                        else
                        {
                            ntimes = Convert.ToInt32(Math.Ceiling((double)newRows / (double)pagination["amountPerPage"]));
                            int dif = (pagination["amountPerPage"] * ntimes) - newRows;
                            progressBar1.Maximum = (pagination["amountPerPage"] * ntimes) - dif;
                        }

                        for (int x = 1; x <= ntimes; x++)
                        {
                            if (x != 1)
                            {
                                Dictionary<string, int> pagination2 = JsonConvert.DeserializeObject<Dictionary<string, int>>(fe.ReadProperty("Pagination", activeSheet.CustomProperties));
                                pagination2["showing"] = pagination2["showing"] + pagination2["amountPerPage"];
                                fe.setProperty("Pagination", activeSheet.CustomProperties, JsonConvert.SerializeObject(pagination2));
                                fe.setProperty("doSomething", activeSheet.CustomProperties, JsonConvert.SerializeObject("false"));
                                await traerDatosDeTabla();
                            }
                            else
                            {
                                pagination["showing"] = 0;
                                pagination["allRows"] = 0;
                                fe.setProperty("Pagination", activeSheet.CustomProperties, JsonConvert.SerializeObject(pagination));
                                fe.setProperty("doSomething", activeSheet.CustomProperties, JsonConvert.SerializeObject("false"));
                                await traerDatosDeTabla();
                            }
                        }
                    }
                }
            }

            label2.ForeColor = Color.Green;
            button2.Enabled = true;
            Stop();
        }
        
        private void downloadData_Load(object sender, EventArgs e)
        {
            this.Activated += AfterLoading;
        }

        private void AfterLoading(object sender, EventArgs e)
        {
            Start();
        }

        public async Task traerDatosDeTabla()
        {
            try
            {
            
                Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);            
                activeSheet.Unprotect("123");

                List<string> columnasDeTabla = JsonConvert.DeserializeObject<List<string>>(fe.ReadProperty("columnasDeTabla", activeSheet.CustomProperties));
                conexionesOData conexionSeleccionada = JsonConvert.DeserializeObject<conexionesOData>(fe.ReadProperty("conexionSeleccionada", activeSheet.CustomProperties));
                conexionTabla conexiontabla = JsonConvert.DeserializeObject<conexionTabla>(fe.ReadProperty("conexionTabla", activeSheet.CustomProperties));
                Dictionary<string, int> pagination = JsonConvert.DeserializeObject<Dictionary<string, int>>(fe.ReadProperty("Pagination", activeSheet.CustomProperties));
                List<List<string>> listaDeDatosCrudosNew = JsonConvert.DeserializeObject<List<List<string>>>(fe.ReadProperty("listaDeDatosCrudosNew", activeSheet.CustomProperties));
                List<int> listOfIdOfDatetime = JsonConvert.DeserializeObject<List<int>>(fe.ReadProperty("ListaDeDatetime", activeSheet.CustomProperties));
                List<int> llavesPrimarias = JsonConvert.DeserializeObject<List<int>>(fe.ReadProperty("idLlavePrimaria", activeSheet.CustomProperties));
                string trigger = JsonConvert.DeserializeObject<string>(fe.ReadProperty("doSomething", activeSheet.CustomProperties));

                //Conexion with OData
                ODataClient client = cnttps.conectWithCredentials(conexionSeleccionada);

                //Create the dictionary for the information
                IEnumerable<IDictionary<string, object>> informacion = Enumerable.Empty<Dictionary<string, object>>();

                //We need know the maximum value of fields in the OData server for this table
                if (pagination["allRows"] != 0)
                {
                    informacion = await client.For(conexiontabla.Tabla).Top(pagination["amountPerPage"]).Skip(pagination["showing"]).FindEntriesAsync();
                }
                else
                {
                    informacion = await client.For(conexiontabla.Tabla).Top(pagination["amountPerPage"]).FindEntriesAsync(annotations);
                    if (pagination["whoCallsTheFunction"] == 0)
                    {
                        progressBar1.Maximum = informacion.Count();
                    }
                    pagination["allRows"] = Convert.ToInt32(annotations.Count);
                    fe.setProperty("Pagination", activeSheet.CustomProperties, JsonConvert.SerializeObject(pagination));
                }

                //Start getting the information
                var datosEnArray = new object[informacion.Count() + 1, columnasDeTabla.Count];
                var datosDeColumnas = new object[0, columnasDeTabla.Count];
                int indexFila = pagination["showing"] + 1;
                int indexColumna = 1;
                if (indexFila == 1)
                {
                    foreach (var col in columnasDeTabla)
                    {
                        datosEnArray[indexFila - 1, indexColumna - 1] = col;
                        indexColumna = indexColumna + 1;
                    }
                }
                else
                {
                    datosEnArray = new object[informacion.Count(), columnasDeTabla.Count];
                    indexFila = 0;
                    pagination["showing"] = pagination["showing"] + 1;
                }

                indexFila = indexFila + 1;
                foreach (var dato in informacion)
                {
                    indexColumna = 1;
                    List<string> datos_crudos = new List<string>();
                    for (int x = 0; x <= columnasDeTabla.Count - 1; x++)
                    {
                        var data_ = "";
                        if (dato.ElementAt(x).Value != null)
                        {
                            data_ = dato.ElementAt(x).Value.ToString();
                            DateTime data_covert;
                            var formats = new[] { "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd" };
                            if (DateTime.TryParseExact(data_, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out data_covert))
                            {
                                if (listOfIdOfDatetime.Any(y => y == indexColumna) == false)
                                {
                                    listOfIdOfDatetime.Add(indexColumna);
                                }
                                data_ = data_covert.ToString("yyyy-MM-dd HH:mm:ss");
                                activeSheet.Cells[indexFila, indexColumna].NumberFormat = "yyyy-MM-dd HH:mm:ss";
                            }
                        }
                        else
                        {
                            data_ = null;
                        }

                        datosEnArray[indexFila - 1, indexColumna - 1] = data_;
                        datos_crudos.Add(data_);
                        indexColumna = indexColumna + 1;
                    }
                    listaDeDatosCrudosNew.Add(datos_crudos);
                    indexFila = indexFila + 1;
                    RefreshProgress(1);
                }

                label1.ForeColor = Color.Green;

                fe.setProperty("listaDeDatosCrudosNew", activeSheet.CustomProperties, JsonConvert.SerializeObject(listaDeDatosCrudosNew));
                fe.setProperty("ListaDeDatetime", activeSheet.CustomProperties, JsonConvert.SerializeObject(listOfIdOfDatetime));

                var firstCell = activeSheet.Cells[pagination["showing"] + 1, 1];
                if (pagination["showing"] != 0)
                {
                    pagination["showing"] = pagination["showing"] - 1;
                }
                var lastCell = activeSheet.Cells[pagination["showing"] + informacion.Count() + 1, columnasDeTabla.Count];

                var range = activeSheet.Range[firstCell, lastCell];
                range.Value2 = datosEnArray;
                range.Font.Color = Color.FromArgb(0, 0, 0);
                range.Borders.Color = Color.FromArgb(163, 163, 163);
                range.Locked = false;

                var first = activeSheet.Cells[1, 1];
                var range2 = activeSheet.Range[first, lastCell];
                range2.Columns.AutoFit();

                cnt.DarFormatoALlavesPrimarias(llavesPrimarias, activeSheet, listaDeDatosCrudosNew.Count);

                if (pagination["showing"] == 0)
                {
                    var list = activeSheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, range, Type.Missing, XlYesNoGuess.xlYes, Type.Missing);
                    list.Name = conexiontabla.Tabla;
                    list.TableStyle = "TableStyleMedium7";

                    var lastCellColumns = activeSheet.Cells[1, columnasDeTabla.Count];
                    var range3 = activeSheet.Range[firstCell, lastCellColumns];
                    range3.Locked = true;
                }

                Globals.Ribbons.Ribbon1.RibbonUI.ActivateTabMso("TabAddIns");

                var focus = activeSheet.get_Range("A1", "A1");
                focus.Select();
                activeSheet.Protect("123");

            }
            catch (WebRequestException exception)
            {
                //CORREGIR ESTA EXCEPCION PORQUE NO DA EL MENSAJE QUE GENERA.
                //StreamReader reader = new StreamReader(exception.Response);
                //MessageBox.Show(readStream.ReadToEnd());
                MessageBox.Show("Access denied");
                this.Close();
            }
            catch (Exception e)
            {
                Worksheet worksheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
                worksheet.Delete();

                MessageBox.Show(e.Message);
                this.Close();
            }
            
        }



        public async void obtenerColumnasDeTabla()
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.Application.ActiveSheet);
            conexionTabla conexiontabla = JsonConvert.DeserializeObject<conexionTabla>(fe.ReadProperty("conexionTabla", activeSheet.CustomProperties));
            //Inicialization of the list for the columns
            List<string> columnasDeTabla = new List<string>();
            int endCol = 0;
            conexionesOData conexionSeleccionada = JsonConvert.DeserializeObject<conexionesOData>(fe.ReadProperty("conexionSeleccionada", activeSheet.CustomProperties));
            ODataClient client = cnttps.conectWithCredentials(conexionSeleccionada);
            //Get the metadata of the OData server
            IEdmModel metadata = await client.GetMetadataAsync<IEdmModel>();
            //Filtering the metadata by the tabla selected
            var entityTypes = metadata.SchemaElements.Where(x => x.Name == conexiontabla.Tabla).OfType<IEdmEntityType>().ToArray();
            //Array for primary keys
            List<int> llavesPrimarias = new List<int>();
            //Cycle for get the differents columns
            foreach (var tabla in entityTypes)
            {
                var columnas = tabla.DeclaredProperties.ToArray();
                progressBar2.Maximum = columnas.Count();
                string[] nombresDeLlavesPrimarias = tabla.DeclaredKey.Select(x => x.Name).ToArray();
                var cantidad = columnas.Count();

                for (int x = 0; x <= cantidad - 1; x++)
                {
                    var columna = columnas.ElementAt(x);
                    if (columna.PropertyKind.ToString() == "Structural")
                    {
                        columnasDeTabla.Insert(x, columna.Name);
                    }

                    if (nombresDeLlavesPrimarias.Contains(columna.Name))
                    {
                        llavesPrimarias.Add(x);
                    }
                    RefreshProgress(2);
                }

                endCol = columnasDeTabla.Count;
            }

            if (fe.ReadProperty("columnasDeTabla", activeSheet.CustomProperties) == null)
            {
                activeSheet.CustomProperties.Add("columnasDeTabla", JsonConvert.SerializeObject(columnasDeTabla));
                activeSheet.CustomProperties.Add("idLlavePrimaria", JsonConvert.SerializeObject(llavesPrimarias));
                activeSheet.CustomProperties.Add("endCol", endCol.ToString());
            }
            else
            {
                fe.setProperty("columnasDeTabla", activeSheet.CustomProperties, JsonConvert.SerializeObject(columnasDeTabla));
                fe.setProperty("idLlavePrimaria", activeSheet.CustomProperties, JsonConvert.SerializeObject(llavesPrimarias));
                fe.setProperty("endCol", activeSheet.CustomProperties, endCol.ToString());
            }

            label3.ForeColor = Color.Green;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
