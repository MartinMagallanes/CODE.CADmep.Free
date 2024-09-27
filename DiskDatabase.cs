using Autodesk.Fabrication;
using Autodesk.Fabrication.Content;
using System.Collections.Generic;
using System.IO;

namespace CODE.CADmep.Free
{
    public static class DiskDatabase
    {
        public static string ItemsPath = Path.GetFullPath(Autodesk.Fabrication.ApplicationServices.Application.ItemContentPath);
        public static IEnumerable<Item> GetDatabaseItems()
        {
            foreach (string str in Directory.EnumerateFiles(DiskDatabase.ItemsPath, "*.itm", SearchOption.AllDirectories))
            //foreach (string str in Directory.EnumerateFiles(DatabaseItemsPath, "*.itm", SearchOption.AllDirectories))
            {
                Item itm = ContentManager.LoadItem(str);
                if (itm != null)
                {
                    yield return itm;
                }
            }
        }
    }
}
