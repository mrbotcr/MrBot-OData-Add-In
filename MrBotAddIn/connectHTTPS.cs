using Simple.OData.Client;
using System;
using System.Net;

namespace MrBotAddIn
{
    public class connectHTTPS
    {
        public ODataClient conectWithCredentials(conexionesOData conexionSeleccionada)
        {
            ODataClientSettings odcSettings = new ODataClientSettings();
            //Define the URL
            Uri uriOdata = new Uri(conexionSeleccionada.Url);
            odcSettings.BaseUri = uriOdata;
            odcSettings.Credentials = new NetworkCredential(conexionSeleccionada.Username, conexionSeleccionada.Password);
            odcSettings.BeforeRequest = requestMessage =>
            {
                requestMessage.Headers.Accept.Clear();
                requestMessage.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                requestMessage.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/xml"));
            };
            odcSettings.WebRequestExceptionMessageSource = new WebRequestExceptionMessageSource();
            ODataClient client = new ODataClient(odcSettings);

            return client;
        }
    }
}
