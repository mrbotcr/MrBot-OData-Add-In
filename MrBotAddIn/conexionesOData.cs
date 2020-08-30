using System.ComponentModel;

namespace MrBotAddIn
{
    public class conexionesOData
    {
        [Category("Connection")]
        [Description("Name to identify the connection.")]
        [DisplayName("Connection name")]

        public string Name
        {
            get;
            set;
        }

        [Category("Connection")]
        [Description("URL of the OData service.")]
        [DisplayName("URL to the OData service")]

        public string Url
        {
            get;
            set;
        }

        [Category("Connection")]
        [Description("Username.")]
        [DisplayName("User name")]

        public string Username
        {
            get;
            set;
        }

        [Category("Connection")]
        [Description("Password.")]

        public string Password
        {
            get;
            set;
        }

    }
}
