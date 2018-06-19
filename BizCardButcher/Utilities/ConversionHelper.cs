using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BizCardButcher.Utilities
{
    class ConversionHelper
    {
        internal static float MMToDots(float mm)
        {
            float dots = InchesToDots(MMToInches(mm));
            return dots;
        }

        internal static float MMToInches(float mm)
        {
            float inches = mm / 25.4f;
            return inches;
        }

        internal static float InchesToDots(float inches)
        {
            float dots = inches * 72f;
            return dots;
        }
    }
}
