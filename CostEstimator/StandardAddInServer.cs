using System;
using System.Runtime.InteropServices;
using Inventor;
using Microsoft.Win32;
using InvAddIn;
using System.Drawing;

namespace CostEstimator
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    [GuidAttribute("881b4da5-3499-4d71-824c-cc30e08b9e0b")]
    public class StandardAddInServer : Inventor.ApplicationAddInServer
    {

        // Inventor application object.
        private Inventor.Application m_inventorApplication;
        private ButtonDefinition m_CostEstimator;
        CostUserControl costEstControl = null;

        private static readonly string addInGUID = "881b4da5-3499-4d71-824c-cc30e08b9e0b";

        bool windowVisible = false;

        public StandardAddInServer()
        {
        }

        #region ApplicationAddInServer Members

        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            // This method is called by Inventor when it loads the addin.
            // The AddInSiteObject provides access to the Inventor Application object.
            // The FirstTime flag indicates if the addin is loaded for the first time.

            // Initialize AddIn members.
            m_inventorApplication = addInSiteObject.Application;

            // Create Icon
            ControlDefinitions controlDefs = m_inventorApplication.CommandManager.ControlDefinitions;

            Icon smallPush = InvAddIn.Properties.Resources.costEstimateIcon;
            Icon largePush = InvAddIn.Properties.Resources.costEstimateIcon;

            stdole.IPictureDisp smallCostIcon = PictureDispConverter.ToIPictureDisp(smallPush);
            stdole.IPictureDisp largCostIcon = PictureDispConverter.ToIPictureDisp(largePush);

            m_CostEstimator = controlDefs.AddButtonDefinition("Estimator", "Open Cost Estimator Window.", CommandTypesEnum.kFilePropertyEditCmdType, addInGUID, "Open Cost Estimator Window.", "Cost Estimator", smallCostIcon, largCostIcon);

            if (firstTime)
            {

                try
                {
                    if (m_inventorApplication.UserInterfaceManager.InterfaceStyle == InterfaceStyleEnum.kRibbonInterface)
                    {
                        // Assembly Button
                        Ribbon assemblyRibbon = m_inventorApplication.UserInterfaceManager.Ribbons["Assembly"];
                        RibbonTab toolsTab = assemblyRibbon.RibbonTabs["id_TabTools"];

                        // Part Buttons
                        Ribbon partRibbon = m_inventorApplication.UserInterfaceManager.Ribbons["Part"];
                        RibbonTab toolsPartTab = partRibbon.RibbonTabs["id_TabTools"];
                        //RibbonTab modelTab = partRibbon.RibbonTabs["id_TabModel"];

                        try
                        {
                            // For ribbon interface
                            // This is a new panel that can be made
                            RibbonPanel panel = toolsTab.RibbonPanels.Add("Cost Estimator", "Autodesk:Cost Estimator:Panel1", addInGUID, "", false);
                            //   CommandControl control1 = panel.CommandControls.AddButton(m_PushParametersButton, true, true, "", false);
                            //  CommandControl control2 = panel.CommandControls.AddButton(m_UpdateIlogicButtton, true, true, "", false);
                            CommandControl control1 = panel.CommandControls.AddButton(m_CostEstimator, true, true, "", false);

                            // Child asy pulling from Parent
                            //                  RibbonPanel panel_asm = assemblyTab.RibbonPanels.Add("Assembly to Parts", "Autodesk:Assembly to Parts:panel_asm", addInGUID, "", false);
                            //             CommandControl control2 = panel_asm.CommandControls.AddButton(m_PullFromParents, true, true, "", false);

                            RibbonPanel pane1_part = toolsPartTab.RibbonPanels.Add("Cost Estimator", "Autodesk:Cost Estimator:pane1_part", addInGUID, "", false);
                            CommandControl control4 = pane1_part.CommandControls.AddButton(m_CostEstimator, true, true, "", false);

                            //RibbonPanel panel_model = toolsTab.RibbonPanels.Add("Cost Estimator", "Autodesk:Cost Estimator:pane1_model", addInGUID, "", false);
                            //CommandControl control5 = panel_model.CommandControls.AddButton(m_CostEstimator, true, true, "", false);

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                    else
                    {
                        // For classic interface, possibly incorrect code
                        CommandBar oCommandBar = m_inventorApplication.UserInterfaceManager.CommandBars["AMxAssemblyPanelCmdBar"];
                        oCommandBar.Controls.AddButton(m_CostEstimator, 0);
                        //    oCommandBar.Controls.AddButton(m_PushParametersButton, 0);
                        //oCommandBar.Controls.AddButton(m_UpdateIlogicButtton, 0);
                    }
                }
                catch
                {
                    // For classic interface, possibly incorrect code
                    CommandBar oCommandBar = m_inventorApplication.UserInterfaceManager.CommandBars["AMxAssemblyPanelCmdBar"];
                    oCommandBar.Controls.AddButton(m_CostEstimator, 0);

                    //    oCommandBar.Controls.AddButton(m_PushParametersButton, 0);
                    //oCommandBar.Controls.AddButton(m_UpdateIlogicButtton, 0);
                }
            }

            m_CostEstimator.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(m_CostEstimator_OnExecute);           
            
            // TODO: Add ApplicationAddInServer.Activate implementation.
            // e.g. event initialization, command creation etc.
        }

        public void Deactivate()
        {
            // This method is called by Inventor when the AddIn is unloaded.
            // The AddIn will be unloaded either manually by the user or
            // when the Inventor session is terminated

            // TODO: Add ApplicationAddInServer.Deactivate implementation

            // Release objects.
            m_inventorApplication = null;

            Marshal.ReleaseComObject(m_CostEstimator);
            m_CostEstimator = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void ExecuteCommand(int commandID)
        {
            // Note:this method is now obsolete, you should use the 
            // ControlDefinition functionality for implementing commands.
        }

        public object Automation
        {
            // This property is provided to allow the AddIn to expose an API 
            // of its own to other programs. Typically, this  would be done by
            // implementing the AddIn's API interface in a class and returning 
            // that class object through this property.

            get
            {
                // TODO: Add ApplicationAddInServer.Automation getter implementation
                return null;
            }
        }

        public void m_CostEstimator_OnExecute(NameValueMap Context)
        {
            if (costEstControl == null)
            {
                CreateDockableWindow();
            }
            else
            {
                costEstControl.DocVis(!costEstControl.dockShown);             
            }
        }

        private void CreateDockableWindow()
        {
            costEstControl = new CostUserControl(m_inventorApplication);

            costEstControl.SetDockableWindow(addInGUID);
            costEstControl.DocVis(true);
        }

        #endregion

    }

    public sealed class PictureDispConverter
    {
        [DllImport("OleAut32.dll", EntryPoint = "OleCreatePictureIndirect", ExactSpelling = true, PreserveSig = false)]
        private static extern stdole.IPictureDisp OleCreatePictureIndirect([MarshalAs(UnmanagedType.AsAny)]
            object picdesc, ref Guid iid, [MarshalAs(UnmanagedType.Bool)]
            bool fOwn);


        static Guid iPictureDispGuid = typeof(stdole.IPictureDisp).GUID;

        private sealed class PICTDESC
        {
            private PICTDESC()
            {
            }


            //Picture Types

            public const short PICTYPE_UNINITIALIZED = -1;
            public const short PICTYPE_NONE = 0;
            public const short PICTYPE_BITMAP = 1;
            public const short PICTYPE_METAFILE = 2;
            public const short PICTYPE_ICON = 3;

            public const short PICTYPE_ENHMETAFILE = 4;

            [StructLayout(LayoutKind.Sequential)]
            public class Icon
            {
                internal int cbSizeOfStruct = Marshal.SizeOf(typeof(PICTDESC.Icon));
                internal int picType = PICTDESC.PICTYPE_ICON;
                internal IntPtr hicon = IntPtr.Zero;
                internal int unused1;

                internal int unused2;

                internal Icon(System.Drawing.Icon icon)
                {
                    this.hicon = icon.ToBitmap().GetHicon();
                }
            }


            [StructLayout(LayoutKind.Sequential)]
            public class Bitmap
            {
                internal int cbSizeOfStruct = Marshal.SizeOf(typeof(PICTDESC.Bitmap));
                internal int picType = PICTDESC.PICTYPE_BITMAP;
                internal IntPtr hbitmap = IntPtr.Zero;
                internal IntPtr hpal = IntPtr.Zero;

                internal int unused;

                internal Bitmap(System.Drawing.Bitmap bitmap)
                {
                    this.hbitmap = bitmap.GetHbitmap();
                }
            }
        }


        public static stdole.IPictureDisp ToIPictureDisp(System.Drawing.Icon icon)
        {
            PICTDESC.Icon pictIcon = new PICTDESC.Icon(icon);
            return OleCreatePictureIndirect(pictIcon, ref iPictureDispGuid, true);
        }


        public static stdole.IPictureDisp ToIPictureDisp(System.Drawing.Bitmap bmp)
        {
            PICTDESC.Bitmap pictBmp = new PICTDESC.Bitmap(bmp);
            return OleCreatePictureIndirect(pictBmp, ref iPictureDispGuid, true);
        }
    }
}
