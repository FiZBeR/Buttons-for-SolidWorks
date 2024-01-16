using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace FirtsBottom
{
    /// <summary>
    /// Our SolidWorks taskpane add-in
    /// <summary>
    public class TaskpaneIntegration : ISwAddin
    {

        #region Private Members

        /// <summary>
        /// The cookie to the current instance of SolidWorks we are running inside of
        /// <summary>
        private int mSwCookie;

        /// <summary>
        /// The taskpane view for our add
        /// <summary>
        private TaskpaneView mtaskpaneView;

        /// <summary>
        /// The UI control that is going to be inside the Soliddworks taskpane view
        /// <summary>
        private TaskpaneHostUI mtaskpaneHost;

        /// <summary>
        /// The current instance of the SolidWorks application
        /// <summary>
        private SldWorks mSolidWorksApplication;

        #endregion

        #region  Public Members

        /// <summary>
        /// The unique Id to the taskpane used for registration in COM
        /// <summary>
        public const string SWTASKPANE_PORID = "SolidWorks.BlankAddin.Taskpane";

        #endregion

        #region SolidWokrs Add-in Callbacks
        /// <summary>
        /// Called when SolidWorks has loaded our add-in and wants us to do our connection logic
        /// </summary>
        /// <param name="ThisSW">The current SolidWorks instance</param>
        /// <param name="Cookie">The current SolidWorks cookie Id</param>
        /// <returns></returns>
        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            // Store a reference to the current SolidWorks intance
            mSolidWorksApplication = (SldWorks)ThisSW;

            // Store cookie id
            mSwCookie = Cookie;

            // Setup callback info
            var ok = mSolidWorksApplication.SetAddinCallbackInfo2(0, this, mSwCookie);

            //Create our UI
            LoadUI();

            //return ok
            return true;
        }

        /// <summary>
        /// Called when SolidWorks is about to unload our add-in and wants us to do our disconnection logic
        /// </summary>
        /// <returns></returns>
        public bool DisconnectFromSW()
        {
            //Clean Up our UI
            UnloadUI();

            //return ok
            return true;

        }

        #endregion

        #region Create UI

        /// <summary>
        /// Create our Taskpane and inject our host UI
        /// </summary>
        private void LoadUI()
        {
            // Find location to our taskpane icon
            var imagePath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).Replace(@"file:\", string.Empty), "icons8-dirección-centro-24.png");

            // Create our Taskpane
            mtaskpaneView = mSolidWorksApplication.CreateTaskpaneView2(imagePath, "Conexición Exitosa con SolidWorks");

            // Load our UI into the taskpane
            mtaskpaneHost = (TaskpaneHostUI)mtaskpaneView.AddControl(TaskpaneIntegration.SWTASKPANE_PORID, string.Empty);
        }

        /// <summary>
        /// Cleanup the taskpane when we disconnect/unload
        /// </summary>
        private void UnloadUI()
        {
            mtaskpaneHost = null;

            // Remove taskpane view
            mtaskpaneView.DeleteView();

            // Release COM reference and cleanup memory
            Marshal.ReleaseComObject(mtaskpaneView);

            mtaskpaneView = null;
        }

        #endregion

        #region COM Registration 

        /// <summary>
        /// The COM registration call to add our registry entries to the SolidWorks add-in registry
        /// </summary>
        /// <param name="t"></param>
        [ComRegisterFunction()]
        private static void ComRegister(Type t)
        {
            var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);

            // Create our registry folder for the add-in
            using (var rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(keyPath))
            {
                // Load add-in when SolidWorks opens
                rk.SetValue(null, 1);

                // Set SolidWorks add-in title and description
                rk.SetValue("Title", "Mi SwAddin");
                rk.SetValue("Description", "Conexción Exitosa!");
            }
        }

        /// <summary>
        /// The COM unregister call to remove our custom entries we added in the COM register function
        /// </summary>
        /// <param name="t"></param>
        [ComUnregisterFunction()]
        private static void ComUnregister(Type t)
        {
            var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);

            // Remove our registry entry
            Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree(keyPath);
        }

        #endregion
    }
}
