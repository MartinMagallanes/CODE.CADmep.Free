using Autodesk.AutoCAD.Internal;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Fabrication;
using Autodesk.Fabrication.Content;
using Autodesk.Fabrication.DB;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using CADapp = Autodesk.AutoCAD.ApplicationServices.Application;
using FabDB = Autodesk.Fabrication.DB.Database;

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
    public static class Commands
    {
        [CommandMethod("RevitSupportReport")]
        public static void RevitSupportReport()
        {
            UI.StartTimer();
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("PATNO,ITEM NAME,2016,2017,2018,2019,2020,2021,2022,2023,2024");
            Dictionary<int, string> patterns = new Dictionary<int, string>();
            IEnumerable<Item> items = GetDatabaseItems();
            foreach (Item itm in GetDatabaseItems())
            {
                int patNo = itm.PatternNumber;
                if (!patterns.ContainsKey(patNo))
                {
                    patterns.Add(patNo, $"{patNo.IsSupportedIn(2016)},{patNo.IsSupportedIn(2017)},{patNo.IsSupportedIn(2018)},{patNo.IsSupportedIn(2019)},{patNo.IsSupportedIn(2020)},{patNo.IsSupportedIn(2021)},{patNo.IsSupportedIn(2022)},{patNo.IsSupportedIn(2023)},{patNo.IsSupportedIn(2024)}");
                }
                sb.AppendLine($"{patNo},{itm.GetPath(true)},{patterns[patNo]}");
            }
            string file = Path.Combine(DatabaseItemsPath, "RevitSupportReport.csv");
            using (StreamWriter streamWriter = new StreamWriter(file))
            {
                streamWriter.Write(sb.ToString());
            }
            OpenFileInApp(file);
            UI.Popup($"{items.Count()} items checked for RevitSupportReport.csv. Duration: {UI.StopTimer()}.");
        }
        [CommandMethod("RemoveCommas")]
        public static void RemoveCommas()
        {
            DialogResult res = MessageBox.Show("RemoveCommas can be a destructive operation. Make a backup of your database before proceeding. Do you want to continue?", "RemoveCommas", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (res == DialogResult.No) return;
            int folders = 0;
            int files = 0;
            int buttons = 0;
            int items = 0;
            string current = "";
            try
            {
                foreach (string d in Directory.GetDirectories(DatabaseItemsPath, "*,*", SearchOption.AllDirectories))
                {
                    Directory.Move(d, d.Replace(", ", " ").Replace(",", " "));
                    folders++;
                }
                foreach (string p in Directory.EnumerateFiles(DatabaseItemsPath, "*.*", SearchOption.AllDirectories))
                {
                    if (!Utils.WcMatchEx(p, "*.itm,*.png", true)) continue;
                    if (!p.Contains(",")) continue;
                    current = p;
                    string path = p.Replace(", ", " ").Replace(",", " ");
                    if (!File.Exists(path))
                    {
                        File.Move(p, path);
                        files++;
                    }
                }
                foreach (ServiceTemplate st in FabDB.ServiceTemplates)
                {
                    foreach (ServiceTab tab in st.ServiceTabs)
                    {
                        foreach (ServiceButton btn in tab.ServiceButtons)
                        {
                            foreach (ServiceButtonItem sbi in btn.ServiceButtonItems)
                            {
                                current = Path.GetFullPath(sbi.ItemPath);
                                if (current.Contains(","))
                                {
                                    string newPath = current.Replace(", ", " ").Replace(",", " ");
                                    if (File.Exists(current) && !File.Exists(newPath))
                                    {
                                        File.Move(current, newPath);
                                        items++;
                                    }
                                    sbi.ItemPath = newPath;
                                }
                            }
                            if (btn.Name.Contains(","))
                            {
                                btn.Name = btn.Name.Replace(", ", " ").Replace(",", " ");
                                buttons++;
                            }
                        }
                    }
                }
                FabDB.SaveServices();
                UI.Popup($"Removed commas from {buttons} button names, {items} items in services, and {folders} folders and {files} files in database folders.");
            }
            catch (SystemException ex)
            {
                UI.Princ($"{ex.GetType().Name} occurred while processing {current}. {ex.Message}");
            }
        }
        static string DatabaseItemsPath = Path.GetFullPath(Autodesk.Fabrication.ApplicationServices.Application.ItemContentPath);
        static string GetPath(this Item itm, bool forCSV)
        {
            if (itm == null)
            {
                throw new ArgumentNullException("itm");
            }
            string name = itm.SourceName;
            string itmPath = Path.GetFullPath(itm.FilePath) + name + ".itm";
            return forCSV ? itmPath.Replace(", ", " ").Replace(",", " ") : itmPath;
        }
        static IEnumerable<Item> GetDatabaseItems()
        {
            foreach (string str in Directory.EnumerateFiles(DatabaseItemsPath, "*.itm", SearchOption.AllDirectories))
            //foreach (string str in Directory.EnumerateFiles(DatabaseItemsPath, "*.itm", SearchOption.AllDirectories))
            {
                Item itm = ContentManager.LoadItem(str);
                if (itm != null)
                {
                    yield return itm;
                }
            }
        }
        static void OpenFileInApp(string path)
        {
            try
            {
                using (Process p = Process.Start(@path))
                {
                    p.StartInfo.UseShellExecute = false;
                }
            }
            catch (System.Exception ex)
            {
                UI.Princ($"Failed to open file in default app. {ex.Message}");
            }
        }
    }
    public static class UI
    {
        public static void Princ(object str)
        {
            if (CADapp.DocumentManager.MdiActiveDocument.Editor == null) return;
            CADapp.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n");
            if (str != null)
            {
                CADapp.DocumentManager.MdiActiveDocument.Editor.WriteMessage(AsString(str));
            }
        }
        private static string AsString(object str)
        {
            string res = "";
            if (str != null)
            {
                if (str is IEnumerable<string>)
                {
                    res = Strcat(str as IEnumerable<string>);
                }
                else
                {
                    res = str.ToString();
                }
            }
            return res;
        }
        public static void Popup(object str)
        {
            Princ(AsString(str));
            CADapp.ShowAlertDialog(AsString(str));
        }
        public static string Strcat(IEnumerable<string> strs)
        {
            string s = null;
            foreach (string str in strs)
            {
                s += "\n" + str;
            }
            return s;
        }
        public static string Strcat(params string[] strs)
        {
            string s = null;
            for (int i = 0; i < strs.Length; i++)
            {
                s += strs[i];
            }
            return s;
        }
        public static Stopwatch _stopwatch = null;
        public static void StartTimer()
        {
            if (_stopwatch == null)
            {
                _stopwatch = Stopwatch.StartNew();
            }
            else
            {
                _stopwatch.Restart();
            }
        }
        public static string TimerElapsed()
        {
            if (_stopwatch == null)
            {
                return "Stopwatch has not been initialized.";
            }
            return _stopwatch.Elapsed.ToString();
        }
        public static string StopTimer()
        {
            if (_stopwatch != null)
            {
                var str = TimerElapsed();
                _stopwatch.Stop();
                _stopwatch = null;
                return str;
            }
            else
            {
                return "Stopwatch has not been initialized.";
            }
        }
    }
}