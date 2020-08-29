using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
//using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
//using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using System.Drawing;
using Newtonsoft.Json;
using Microsoft.Office.Tools.Ribbon;

namespace MrBotAddIn
{
    public partial class ThisAddIn
    {
                
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            
            this.Application.WorkbookActivate += Application_WorkbookOpen;
        }

        public void Application_WorkbookOpen(Workbook Doc)
        {

            Sheets sheets = Doc.Worksheets;
            
            connections con = new connections();
            funcionesEspeciales fe = new funcionesEspeciales();
            foreach (Worksheet sheet in sheets)
            {
                if (fe.ReadProperty("conexionSeleccionada", sheet.CustomProperties) != null)
                {
                    DocEvents_ChangeEventHandler EventDel_CellsChange = new DocEvents_ChangeEventHandler(con.WorksheetChangeEventHandler);
                    sheet.Change += EventDel_CellsChange;
                }
            }            
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        #region Código generado por VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
            
        }
        
        #endregion
    }
}
