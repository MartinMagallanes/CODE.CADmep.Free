using System.Collections.Generic;
using System.Diagnostics;
using CADapp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace CODE.CADmep.Free
{
    public static class Sys
    {
        public static void OpenFileInApp(string path)
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
