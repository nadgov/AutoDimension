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
    #region SolidDrawing
    public class SolidDrawing
    {
        #region Fields
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private DrawingDoc swDraw;
        private BalloonOptions BalloonOpt;
        private List<Dimension> DimList;
        private Dictionary<Dimension, List<double>> DimTolDict;

        private List<double[]> _ListDim_Tol;
        #endregion

        #region Constructor
        public SolidDrawing(SldWorks swApp, ModelDoc2 swModel)
        {
            this.swApp = swApp;
            this.swModel = swModel;
            this.swDraw = (DrawingDoc)this.swModel;
            
            this.DimList = GetDimensions();
            this.DimTolDict = GetDimTolDict();
            this._ListDim_Tol = GetListDim_Tol();
            this.BalloonOpt = SetBalloonOptions();

            //CreateBalloons();
            AddDimensionNumber();
            this.swModel.EditRebuild3();

        }
        #endregion

        #region Properties

        public List<double[]> ListDim_Tol { get { return this._ListDim_Tol; } }

        #endregion

        #region Methods
        #region GetDimensions
        private List<Dimension> GetDimensions()
        {
            List<Dimension> DimList = new List<Dimension>();
            View swView = this.swDraw.GetFirstView();
            Dimension swDim;
            if (swView == null)
                throw new AutoDimensionException("Unable to find drawing views");
            while (swView != null)
            {
                var arrAnnotation = swView.GetDisplayDimensions();
                if (arrAnnotation != null)
                {
                    foreach (DisplayDimension DispDim in arrAnnotation)
                    {
                        swDim = (Dimension)DispDim.GetDimension2(0);
                        DimList.Add(swDim);
                    }
                }
                swView = swView.GetNextView();
            }

            return DimList;
        }
        #endregion

        #region GetDimTolDict
        private Dictionary<Dimension, List<double>> GetDimTolDict()
        {
            Dictionary<Dimension, List<double>> dim_tol_dict = new Dictionary<Dimension, List<double>>();
            //returns a list of all tolerances in a drawing in meters
            List<double> Tolerances = new List<double>();
            double MaxVal = 0, MinVal = 0;
            foreach (Dimension dim in this.DimList)
            {
                //Console.WriteLine(dim.Value);
                DimensionTolerance Tolerance = (DimensionTolerance)dim.Tolerance;
                if (Tolerance.GetMaxValue2(out MaxVal) == 0 && Tolerance.GetMinValue2(out MinVal) == 1) // only max value is available (probably symetric tolerance)
                {
                    Tolerances.Add(MaxVal); Tolerances.Add(MaxVal); // Symetric tolerance. Add twice the MaxVal and it will be considered as MinVal to.
                }

                if (Tolerance.GetMaxValue2(out MaxVal) == 0 && Tolerance.GetMinValue2(out MinVal) == 0) // both max value and min value are available (probably non-symetric tolerance)
                {
                    Tolerances.Add(MaxVal); Tolerances.Add(MinVal);
                }

                if (Tolerance.GetMaxValue2(out MaxVal) == 1 && Tolerance.GetMinValue2(out MinVal) == 1) // max value and min value are not available (probably no tolerance)
                {
                    Tolerances.Add(-1); Tolerances.Add(-1); // no tolerance, -1 is considered invalid (General Tolerance)
                    //continue; // need to think what to do with this
                }
                dim_tol_dict.Add(dim, new List<double>());
                dim_tol_dict[dim].AddRange(Tolerances);
                Tolerances.Clear();
            }
            return dim_tol_dict;
        }
        #endregion

        #region GetListDim_Tol
        private List<double[]> GetListDim_Tol()
        {
            ///<summary>
            ///ListDim_Tol is a list containing the dimension and its tolerance. This is the final list needed to be exported to Excel
            ///Format: {dimension value, upper limit tolerance, lower limit tolerance}
            ///Units: millimeter [mm]
            ///</summary>
            List<double[]> ListDim_Tol = new List<double[]>();
            double[] dim_tol = { 0.0, 0.0, 0.0 };
            foreach (Dimension dim in this.DimTolDict.Keys)
            {
                dim_tol[0] = dim.Value; dim_tol[1] = this.DimTolDict[dim][0] * 1000; dim_tol[2] = this.DimTolDict[dim][1] * 1000;
                //Format: dim_tol[0] = dimension value, dim_tol[1] = upper limit tolerance, dim_tol[2] = lower limit tolerance
                ListDim_Tol.Add(dim_tol);
                dim_tol[0] = 0.0; dim_tol[1] = 0.0; dim_tol[2] = 0.0;
            }
            return ListDim_Tol;
        }
        #endregion

        #region SetBalloonOptions
        private BalloonOptions SetBalloonOptions()
        {
            BalloonOptions options = this.swModel.Extension.CreateBalloonOptions();
            options.Style = (int)swBalloonStyle_e.swBS_Circular;
            options.Size = (int)swBalloonFit_e.swBF_Tightest;
            options.UpperTextContent = (int)swBalloonTextContent_e.swBalloonTextCustom;
            options.LowerTextContent = (int)swBalloonTextContent_e.swBalloonTextCustom;
            options.ShowQuantity = false;
            //options.ItemNumberStart = 1;
            //options.ItemNumberIncrement = 1;
            //options.ItemOrder = (int)swBalloonItemNumbersOrder_e.swBalloonItemNumbers_DoNotChangeItemNumbers;
            return options;
        }
        #endregion

        #region EditTextOfBalloon
        private bool EditTextOfBalloon(Note swBalloon, string text)
        {
            if (!swBalloon.IsBomBalloon())
                throw new AutoDimensionException("Not a balloon");
            bool success = swBalloon.SetBomBalloonText((int)swBalloonTextContent_e.swBalloonTextCustom, text, (int)swBalloonTextContent_e.swBalloonTextCustom, "");
            return success;
        }
        #endregion

        #region ChangeBalloonColor
        private void ChangeBalloonColor(Note swBalloon, Color color)
        {
            int COLORREF = ColorTranslator.ToWin32(color);
            Annotation swAnn = (Annotation)swBalloon.GetAnnotation();
            if(swAnn != null)
                swAnn.Color = COLORREF;
        }
        #endregion

        #region EraseBalloonLeader
        private void EraseBalloonLeader(Note swBalloon)
        {
            Annotation swAnn = (Annotation)swBalloon.GetAnnotation();
            if (swAnn != null)
                swAnn.SetLeader3((int)swLeaderStyle_e.swNO_LEADER, (int)swLeaderSide_e.swLS_SMART, false, false, false, false);
        }
        #endregion

        #region ActivateAllViews
        private bool ActivateAllViews()
        {
            View swView = swDraw.GetFirstView();
            while (swView != null)
            {
                bool activate = swDraw.ActivateView(swView.GetName2());
                if (activate == false)
                    return false;
                swView = swView.GetNextView();
            }
            return true;
        }
        #endregion

        #region CreateBalloons
        private void CreateBalloons()
        {
            bool activate = ActivateAllViews();
            if (activate == true)
            {
                this.swModel.ClearSelection2(true);
                for (int i = 0; i < this.DimList.Count; i++)
                {
                    string name = this.DimList[i].GetNameForSelection();
                    string number = (i + 1).ToString();
                    Color color = Color.Red;
                    bool select = this.swModel.Extension.SelectByID2(name, "DIMENSION", 0, 0, 0, false, 0, null, (int)swSelectOption_e.swSelectOptionDefault);
                    if (select != false)
                    {
                        Note swBalloon = this.swModel.Extension.InsertBOMBalloon2(this.BalloonOpt);
                        bool edit = EditTextOfBalloon(swBalloon, number);
                        ChangeBalloonColor(swBalloon, color);
                        EraseBalloonLeader(swBalloon);
                        //swDraw.ForceRebuild();
                    }
                }
            }
        }
        #endregion

        #region AddDimensionNumber
        private void AddDimensionNumber()
        {
            int num = 1;
            DisplayDimension swDispDim;
            View swView = this.swDraw.GetFirstView();
            string CurrSuffix, NewSuffix, Circle = "C#-";
            while (swView != null)
            {
                swDispDim = swView.GetFirstDisplayDimension5();
                while (swDispDim != null)
                {
                    if (swDispDim.IsHoleCallout()) // GetText methid does not support hole callouts
                    {
                        num++;
                        continue;
                    }
                    CurrSuffix = swDispDim.GetText((int)swDimensionTextParts_e.swDimensionTextSuffix);
                    NewSuffix = CurrSuffix + " <" + Circle + num.ToString() + ">";
                    num++;
                    swDispDim.SetText((int)swDimensionTextParts_e.swDimensionTextSuffix, NewSuffix);
                    swDispDim = swDispDim.GetNext5();
                }
                swView = swView.GetNextView();
            }
        }
        #endregion
        #endregion
    }
    #endregion
}
