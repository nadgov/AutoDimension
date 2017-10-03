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
    class Program
    {
        static void Main(string[] args)
        {

            //SolidDrawing SldDrw = new SolidDrawing(new SldWorks());
            try
            {
                AutoDimensionController Controller = new AutoDimensionController();
                Controller.Launch();
                //Controller.ReleaseSolid();
                Console.WriteLine(Controller.SldDrw.ToString());
                Console.ReadKey();
                return;
            }
            catch (AutoDimensionException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
                return;
            }
        }
    }
}
