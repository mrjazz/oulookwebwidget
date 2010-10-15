using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;

namespace OutlookPanel
{
    
    public partial class ThisAddIn
    {

        Microsoft.Office.Tools.CustomTaskPane taskPane;
        TaskPanel browser;
       
        private Office.CommandBar menuBar;
        private Office.CommandBarPopup menuControl;
        private Office.CommandBarButton menuButton;
        Outlook.Explorer explorer;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            explorer = this.Application.ActiveExplorer();
            browser = new TaskPanel();
            taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(browser, "Web browser", explorer);
            taskPane.Visible = true;
            taskPane.Width = 245;
            taskPane.VisibleChanged += new System.EventHandler(ChangeButtonFace);
            taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            taskPane.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;
            AddMenuBar();
        }

        private void ChangeButtonFace(object sender, System.EventArgs e)
        {
            setMenuButtonFace();
        }

        private void setMenuButtonFace()
        {
            if (taskPane.Visible)
            {
                menuButton.FaceId = 3110;
            }
            else
            {
                menuButton.FaceId = 0;
            }
        
        
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Properties.Settings.Default.Save();
            this.CustomTaskPanes.Remove(taskPane);
            this.Application = null;
        }

        private void AddMenuBar()
        {
            try
            {
                menuBar = this.Application.ActiveExplorer().CommandBars.ActiveMenuBar;
                menuControl = (Office.CommandBarPopup)menuBar.Controls[3];
                if (menuControl != null)
                {
                    menuButton = (Office.CommandBarButton)menuControl.Controls. 
                    Add(Office.MsoControlType.msoControlButton, missing,
                        missing, 8, true);
                    menuButton.Style = Office.MsoButtonStyle.
                        msoButtonIconAndCaption;
                    menuButton.Caption = "Web browser";
                    menuButton.Tag = "c123";
                    menuButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(buttonOne_Click);
                    menuControl.Visible = true;
                    setMenuButtonFace();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void buttonOne_Click(Office.CommandBarButton ctrl, ref bool cancel)
        {
            if (taskPane.Visible)
            {
                taskPane.Visible = false;
            }
            else
            {
                taskPane.Visible = true;
            }
            setMenuButtonFace();
        }
        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion

    }
}
