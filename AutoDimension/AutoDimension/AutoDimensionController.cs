#region using
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using SolidWorks;
using Microsoft.VisualBasic;
using SolidWorks.Interop.swconst;
using System.Collections;
using SolidWorks.Interop.sldworks;
using System.Runtime.InteropServices;
using System.Drawing;
//using Newtonsoft.Json;
#endregion
namespace AutoDimension
{
    #region AutoDimensionController
    public class AutoDimensionController
    {
        #region Fields
        private SldWorks _swApp;
        private ModelDoc2 _swModel;
        private int _ActiveDocType;
        private SolidDrawing _SldDrw;
        #endregion

        #region Constructor
        public AutoDimensionController()
        {
            if (this._swApp == null)
                this._swApp = new SldWorks();
            if (this._swApp == null)
                throw new AutoDimensionException("Fatal Error: Unable to connect to SolidWorks");
            this._swModel = this._swApp.ActiveDoc;
            if (this._swModel == null)
                throw new AutoDimensionException("Fatal Error: Unable to get active document");
            this._ActiveDocType = this._swModel.GetType();
        }
        #endregion

        #region Properties

        public SldWorks swApp { get { return this._swApp; } }
        public ModelDoc2 swModel { get { return this._swModel; } }
        public SolidDrawing SldDrw { get { return this._SldDrw; } }
        #endregion

        #region Methods

        #region Launch
        public void Launch()
        {
            SolidDrawing SldDrw;
            if (this._ActiveDocType != (int)swDocumentTypes_e.swDocDRAWING)
                throw new AutoDimensionException("This macro only works on drawings");
            SldDrw = new SolidDrawing(this._swApp, this._swModel);
            this._SldDrw = SldDrw;
        }
        #endregion

        #region ReleaseSolid
        public void ReleaseSolid()
        {
            this._swApp = null;
        }
        #endregion

        #endregion
    }
    #endregion
}
