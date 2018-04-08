using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AutoScoreFiller {
    class ResourceCulture {

        public static void SetCurrentCulture(string name) {
            if (string.IsNullOrEmpty(name)) {
                name = "zh-CN";
            }
            Thread.CurrentThread.CurrentCulture = new CultureInfo(name);
        }

        public static string GetString(string id) {
            string strCurLang = "";
            try {
                ResourceManager rm = new ResourceManager("AutoScoreFiller.Language", Assembly.GetExecutingAssembly());
                CultureInfo ci = Thread.CurrentThread.CurrentCulture;
                strCurLang = rm.GetString(id, ci);
            }
            catch {
                strCurLang = "No id:" + id + ", please add.";
            }
            return strCurLang;
        }
    }
}

