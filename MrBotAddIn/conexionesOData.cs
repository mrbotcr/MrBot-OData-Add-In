using System.ComponentModel;

namespace MrBotAddIn
{
    public class conexionesOData
    {
        [Category("Connection")]
        [Description("Url of the OData server.")]

        public string Url
        {
            get;
            set;
        }

        [Category("Connection")]
        [Description("Name to identify the connection.")]

        public string Name
        {
            get;
            set;
        }

        [Category("Connection")]
        [Description("Username.")]

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
