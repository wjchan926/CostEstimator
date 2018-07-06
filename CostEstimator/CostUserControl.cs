using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Inventor;
using System.Runtime.InteropServices;

namespace InvAddIn
{
    public partial class CostUserControl : UserControl
    {
        Inventor.Application invApp;

        DockableWindow oWindow;
        int dockHeight = 500;
        int dockWidth = 500;
        public bool dockShown { get; private set; } = false;

        int initialDockHeight = 500;
        int initialDockWidth = 500;
        // bool inAction = false;

        string addInCLSIDStr;
        CostForm costForm;

        public CostUserControl(Inventor.Application currentInv)
        {
            invApp = currentInv;
            InitializeComponent();
        }

        private void CostUserControl_Load(object sender, EventArgs e)
        {

        }



        public void SetDockableWindow(string addinCLS)
        {
            UserInterfaceManager oUserInterfaceMgr = invApp.UserInterfaceManager;

            addInCLSIDStr = "{" + addinCLS + "}";

            try
            {
                oWindow = oUserInterfaceMgr.DockableWindows.Add(addInCLSIDStr, "CostEstimatorWindow", "Cost Estimator");
                oWindow.AddChild(CreateChildDialog());
                oWindow.DisabledDockingStates = DockingStateEnum.kDockTop & DockingStateEnum.kDockBottom;
                oWindow.DockingState = DockingStateEnum.kDockRight;
                oWindow.Visible = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public long CreateChildDialog()
        {
            if (costForm != null)
            {
                costForm.Dispose();
                costForm = null;
            }

            costForm = new CostForm(invApp);
            costForm.Show(new WindowWrapper((IntPtr)invApp.MainFrameHWND));
            return costForm.Handle.ToInt64();            
        }

        public void DocVis(bool visible)
        {
            UserInterfaceManager oUserInterfaceMgr = invApp.UserInterfaceManager;
            oUserInterfaceMgr.ShowBrowser = true;

            if (visible)
            {
  
                oWindow.Visible = visible;
                if (!dockShown)
                {
                    oWindow.DockingState = DockingStateEnum.kDockRight;
                    oWindow.SetMinimumSize(dockHeight, dockWidth);
                    oWindow.Width = initialDockWidth;
                    oWindow.Height = initialDockHeight;
                    dockShown = true;
                }
            }
            else
            {
                dockShown = false;
                oWindow.Visible = false; 
            }


        }


    }

    public class WindowWrapper : IWin32Window
    {
        private IntPtr _hwnd;

        public WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }

        public IntPtr Handle
        {
            get { return _hwnd; }
        }        
    }
}
