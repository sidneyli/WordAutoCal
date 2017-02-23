using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Globalization;

namespace AutoCal
{
    /// <summary>
    /// Represents the AutoCal addin class.
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary>
        /// Addin startup event handler.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event Arguments.</param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Link the window select change event with handler.
            this.Application.WindowSelectionChange +=
                new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
        }

        /// <summary>
        /// Addin shutdown event handler.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event Arguments.</param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.Application.DisplayStatusBar = false;
        }

        /// <summary>
        /// Window select change event handler.
        /// </summary>
        /// <param name="selection">The selection.</param>
        void Application_WindowSelectionChange(Word.Selection selection)
        {
            // Initialize a summation
            double summation = 0;

            // Return early if the selection is not a collection of columns
            if (!Word.WdSelectionType.wdSelectionColumn.Equals(selection.Type)
                && !Word.WdSelectionType.wdSelectionRow.Equals(selection.Type))        
                return;

            // Iterate through the cells to get summation
            foreach (Word.Cell cell in selection.Cells)
            {
                double temp = 0;

                // Trim the tailing format chars
                if (double.TryParse(cell.Range.Text.TrimEnd(new char[] { '\r', '\a' }), out temp))
                    summation += temp;
            }

            // Change the status bar
            this.Application.StatusBar = $"The summation of selected numbers is {summation.ToString("N", CultureInfo.InvariantCulture)}";
        }
           
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
