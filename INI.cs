using Autodesk.AutoCAD.Runtime;
using System;
using System.Linq;
using System.Reflection;

namespace CODE.CADmep.Free
{
    public class INI : IExtensionApplication
    {
        void IExtensionApplication.Initialize()
        {
            var acadpath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
            string acadYear = new String(acadpath.ToCharArray().Where(p => char.IsDigit(p) == true).ToArray());
            try
            {
                Assembly.LoadFrom($@"C:\Program Files\Autodesk\Fabrication {acadYear}\CADmep\FabricationAPI.dll");
            }
            catch { }
            UI.Princ("\nLoaded CODE.CADmep.Free!");
        }
        void IExtensionApplication.Terminate() { }
    }
}