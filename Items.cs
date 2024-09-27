using Autodesk.AutoCAD.Internal;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Fabrication;
using Autodesk.Fabrication.DB;
using System;
using System.IO;
using System.Windows.Forms;
using FabDB = Autodesk.Fabrication.DB.Database;

namespace CODE.CADmep.Free
{
    public static class Items
    {
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
                foreach (string d in Directory.GetDirectories(DiskDatabase.ItemsPath, "*,*", SearchOption.AllDirectories))
                {
                    Directory.Move(d, d.Replace(", ", " ").Replace(",", " "));
                    folders++;
                }
                foreach (string p in Directory.EnumerateFiles(DiskDatabase.ItemsPath, "*.*", SearchOption.AllDirectories))
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
        public static string GetPath(this Item itm, bool forCSV)
        {
            if (itm == null)
            {
                throw new ArgumentNullException("itm");
            }
            string name = itm.SourceName;
            string itmPath = Path.GetFullPath(itm.FilePath) + name + ".itm";
            return forCSV ? itmPath.Replace(", ", " ").Replace(",", " ") : itmPath;
        }
    }
}
