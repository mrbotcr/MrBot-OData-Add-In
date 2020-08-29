using Simple.OData.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Data.Edm;
using Newtonsoft.Json;
using System.Net;
using System.Drawing;

namespace MrBotAddIn
{
    public partial class nuevaConexion : Form
    {
        public Ribbon1 ribbon = new Ribbon1();
        
        conexionesOData datosDeConexion = new conexionesOData();

        public nuevaConexion()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Owner.Enabled = false;
            label1.Text = "";
            propertyGrid1.SelectedObject = datosDeConexion;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (datosDeConexion.Name != "")
                {
                    /*
                        In the properties of the complement we are storing the connections, 
                        then we obtain the existing connections to add the new one.
                        If we do not have a connection we establish that we will create a list of connections.
                        We validate that there is no connection with a duplicate name and 
                        add it to the list of connections, write it in the properties of the project and save the information.
                    */
                    List<conexionesOData> lista = JsonConvert.DeserializeObject<List<conexionesOData>>(MrBotAddIn.Properties.Settings.Default.jsonDeConexiones);
                    if (lista == null)
                        lista = new List<conexionesOData>();
                    if (lista.Where(x => x.Name == datosDeConexion.Name).Count() == 0)
                    {
                        lista.Add(datosDeConexion);
                        MrBotAddIn.Properties.Settings.Default.jsonDeConexiones = JsonConvert.SerializeObject(lista);
                        MrBotAddIn.Properties.Settings.Default.Save();

                        //connections connections = new connections();
                        //connections.ShowDialog();
                        this.Owner.Enabled = true;
                        this.Close();
                        
                    }
                    else
                    {
                        MessageBox.Show("You already have a connection with this name.");
                    }
                }
                else
                {
                    MessageBox.Show("Enter a new name for the connection.");
                }
            }catch
            {
                MessageBox.Show("The connection could not be created, check the provided URL.");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Owner.Enabled = true;
            this.Close();
        }

        private void propertyGrid1_PropertyValueChanged(object s, PropertyValueChangedEventArgs e)
        {
            button2.Enabled = false;
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            
            try
            {
                label1.Text = "Connecting......";
                label1.ForeColor = Color.Black;
                button2.Enabled = false;
                if (datosDeConexion.Name != null & datosDeConexion.Url != null )
                {
                    /* 
                        We create an Odata Client Settings to define the URL 
                        with which we establish the connection
                    */
                    ODataClientSettings odcSettings = new ODataClientSettings();
                    //Define the URL
                    Uri uriOdata = new Uri(datosDeConexion.Url);
                    odcSettings.BaseUri = uriOdata;
                    odcSettings.Credentials = new NetworkCredential(datosDeConexion.Username, datosDeConexion.Password);
                    odcSettings.BeforeRequest = requestMessage =>
                    {
                        requestMessage.Headers.Accept.Clear();
                        requestMessage.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                        requestMessage.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/xml"));
                    };

                    /*
                    We establish the connection and we bring the metadata to know if it is connected
                    */
                    ODataClient client = new ODataClient(odcSettings);

                    IEdmModel metadata = await client.GetMetadataAsync<IEdmModel>();
                    var entityTypes = metadata.SchemaElements.OfType<IEdmEntityType>().ToArray();
                    label1.Text = "Successful connection...";
                    label1.ForeColor = Color.Green;
                    button2.Enabled = true;
                }
                else
                {
                    label1.Text = "";
                    label1.ForeColor = Color.Black;
                    button2.Enabled = false;
                    MessageBox.Show("Please complete all the information.");
                }
            }catch(Exception ex)
            {
                label1.Text = "Connection not established...";
                label1.ForeColor = Color.Red;
            }
        }
    }

}
