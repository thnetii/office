using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace THNETII.Office.Office2Git.WordPlugin
{
    public partial class Office2GitWordAddin
    {
        private void WordAddinStartup(object sender, System.EventArgs e)
        {
        }

        private void WordAddinnShutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(WordAddinStartup);
            this.Shutdown += new System.EventHandler(WordAddinnShutdown);
        }
        
        #endregion
    }
}
