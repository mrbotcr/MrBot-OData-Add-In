using Microsoft.Office.Interop.Excel;

namespace MrBotAddIn
{
    public class funcionesEspeciales
    {

        public string ReadProperty(string propertyName, CustomProperties properties)
        {

            foreach (CustomProperty prop in properties)
            {
                if (prop.Name == propertyName)
                {
                    return prop.Value;
                }
            }
            return null;
        }

        public void setProperty(string propertyName, CustomProperties properties, object valor)
        {
            foreach (CustomProperty prop in properties)
            {
                if (prop.Name == propertyName)
                {
                    prop.Value = valor;
                }
            }
        }
    }
}
