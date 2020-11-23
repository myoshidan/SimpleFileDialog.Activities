using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleFileDialog.Enums
{
    [Flags]
    public enum FileTypes
    {
        All = 0,
        Excel = 1,
        CSV = 2,
        PDF = 4,
        Text = 8,
        PowerPoint = 16,
        Word = 32
    }
}
