using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointQuestionnaire.Services
{
    public static class AppWapper
    {
        public static PowerPoint.Application App { get; private set; }

        public static void SetApp(PowerPoint.Application application)
        {
            App = application;
        }
    }
}
