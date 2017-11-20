using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
//using ExcelDna.Utilities;
using System.Diagnostics;

namespace ThermoExcel
{

    public class ExcelInt
    {
        // ************************************************************************************************
        // All inputs from Excel are 2D
        // I am making all outputs to Excel 2D so that vectors can be outputted as either columns or rows
        // *************************************************************************************************

        /* public class MyFunctions
         {
             // This function returns a string that describes its argument.
             // For arguments defined as object type, this shows all the possible types that may be received.
             // Also try this function after changing the 
             // [ExcelArgument(AllowReference=true)] attribute.
             // In that case we allow references to be passed (registerd as type R). 
             // By default the function will be registered not
             // to receive references AllowReference=false (type P).
             [ExcelFunction(Description = "Describes the value passed to the function.", IsMacroType = true)]
             public static string Describe([ExcelArgument(AllowReference = false)]object arg)
             {
                 if (arg is double)
                     return "Double: " + (double)arg;
                 else if (arg is string)
                     return "String: " + (string)arg;
                 else if (arg is bool)
                     return "Boolean: " + (bool)arg;
                 else if (arg is ExcelError)
                     return "ExcelError: " + arg.ToString();
                 else if (arg is object[,])
                     // The object array returned here may contain a mixture of different types,
                     // reflecting the different cell contents.
                     return string.Format("Array[{0},{1}]", ((object[,])arg).GetLength(0), ((object[,])arg).GetLength(1));
                 else if (arg is ExcelMissing)
                     return "Missing";
                 else if (arg is ExcelEmpty)
                     return "Empty";
                 else if (arg is ExcelReference)
                     return "Reference: " + XlCall.Excel(XlCall.xlfReftext, arg, true);
                 else
                     return "!?Unheard Of";
             }*/

        internal static bool IsEmpty(object arg)
        {
            if (arg is ExcelEmpty)
                return true;
            else
                return false;
        }

        internal static bool IsMissing(object arg)
        {
            if (arg is ExcelMissing)
                return true;
            else
                return false;
        }

        [ExcelFunction(IsMacroType = true)]
        internal static string GetWorkBookName()
        {
            object reference = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
            string sheetName = (string)XlCall.Excel(XlCall.xlSheetNm, reference);
            string nameLessBrak = sheetName.TrimStart((char)91);
            int locPeriod = nameLessBrak.IndexOf(".", 0);
            string name;
            if (locPeriod < 0)
                name = nameLessBrak;
            else
                name = nameLessBrak.Remove(locPeriod);

            return name;
        }

        public class MyArray
        {
            public int nRows;
            public int nCols;
            public int n1D;
            public bool is1DArray = false;
            public bool is2DArray = false;
            public bool isRowArray = false;
            public bool isColArray = false;
            public object[,] array;
            public MyArray(object[,] myA)
            {
                array = myA;
                nRows = array.GetLength(0);
                nCols = array.GetLength(1);
                if (nRows > 1 && nCols > 1)
                    is2DArray = true;

                is1DArray = true;

                if (nCols > 1)
                {
                    n1D = nCols;
                    isRowArray = true;
                }
                else
                {
                    n1D = nRows;
                    isColArray = true;// default for 1 cell array
                }
            }
        }

        //public object[] oneDToCSArray
        //{// returns a C# one d array
        //    object[] ret = new object[]

        //}
        public object[,] oneDToColArray(object[] ar)
        {// use for One D inputs
            object[,] ret = new object[ar.Length, 1];
            for (int i = 0; i < ar.Length; i++)
                ret[i, 0] = ar[i];
            return ret;
        }

        public object[,] oneDToRowArray(object[] ar)
        {// use for One D inputs
            object[,] ret = new object[1, ar.Length];
            for (int i = 0; i < ar.Length; i++)
                ret[0, i] = ar[i];
            return ret;
        }

        public object[,] oneDToColArray(int Length)
        {// use with inputs
            object[,] ret = new object[Length, 1];
            return ret;
        }

        public object[,] oneDToRowArray(int Length)
        {// use with inputs
            object[,] ret = new object[Length, 1];
            return ret;
        }

        public object[,] retColArray(object[] ar)
        {
            object[,] ret = new object[ar.Length, 1];
            for (int i = 0; i < ar.Length; i++)
                ret[i, 0] = ar[i];
            return ret;
        }

        public object[,] retRowArray(object[] ar)
        {
            object[,] ret = new object[1, ar.Length];
            for (int i = 0; i < ar.Length; i++)
                ret[0, i] = ar[i];
            return ret;
        }

        public object[,] returnMessage(string message)
        {
            object[,] ret = new object[1, 1];
            ret[0, 0] = message;
            return ret;
        }

        public object[,] returnExceptionMessage(string message, Exception e)
        {
            object[,] ret = new object[1, 1];
            ret[0, 0] = message + e.Message;
            return ret;
        }

    }// end ExcelInt

}// end namespace CBExcel