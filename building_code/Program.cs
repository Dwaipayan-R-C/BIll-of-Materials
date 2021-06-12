using System;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Runtime.InteropServices;


namespace building_code
{
    class Program
    {
        static void Main ()
        {
            SldWorks swApp = (SldWorks) Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
            ModelDoc2 swModelDrawing = (ModelDoc2) swApp.ActiveDoc;
            Console.WriteLine(swModelDrawing.GetType());            

        }
    }
}
