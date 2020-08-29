using System.Collections.Generic;

namespace MrBotAddIn
{
    public class changeInformation
    {
        public int ch_idRow
        {
            get;
            set;
        }

        public int ch_idRowExcel
        {
            get;
            set;
        }
        public string ch_action
        {
            get;
            set;
        }

        public object ch_newValue
        {
            get;
            set;
        }

        public int ch_columnId
        {
            get;
            set;
        }

        public string ch_column
        {
            get;
            set;
        }

        public List<llavesPrimariasClass> ch_llaves_primarias
        {
            get;
            set;
        }
    }
}
