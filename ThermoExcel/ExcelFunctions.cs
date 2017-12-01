using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Thermodynamics;

using Therm = Thermodynamics.Thermo;
using ML = Monolith;

namespace ThermoExcel
{
    public static class MyFunctions
    {
        public static ExcelInt myExcelInt = new ExcelInt();// seem to need this so that interface functions are not exported

        [ExcelFunction(Description = "Lists the type of units that can be converted to units of the same or similar types")]
        public static object[,] UnitTypesList_([ExcelArgument("Enter true for a column list")]bool isCol = false)
        {

            List<string> ListOfTypes = ML.CallHandler.ListUnitTypes();
            string[] myTypes = ListOfTypes.ToArray();

            object[,] ret;
            if (isCol)
                ret = myExcelInt.oneDToColArray(myTypes);
            else
                ret = myExcelInt.oneDToRowArray(myTypes);

            return ret;
        }

        [ExcelFunction(Description = "Converts a number from an input unit to that of an output unit of the same type")]
        public static object UnitConvert_([ExcelArgument("Value to be converted")]     double valueIn,
                                            [ExcelArgument("Original units")]            string inUnit,
                                            [ExcelArgument("Desired units")]            string outUnit,
                                            [ExcelArgument("Mol Wt, Density or valid gas or feedstock name")] object mwOrGasName)
        {
            return ML.CallHandler.ConvertUnits(inUnit, outUnit, valueIn, mwOrGasName);
        }

        [ExcelFunction(Description = "Lists the valid units for a specific type of unit")]
        public static object[,] UnitsList_([ExcelArgument("Enter a valid unit type")]string unitType, [ExcelArgument("Enter true for a column list")]bool isCol = false)
        {
            string[] myString = ML.CallHandler.ListUnits(unitType);

            object[,] ret;
            if (isCol)
                ret = myExcelInt.oneDToColArray(myString);
            else
                ret = myExcelInt.oneDToRowArray(myString);

            return ret;
        }

        [ExcelFunction(Description = "Lists the species in Thermo Database")]
        public static object[,] SpeciesList_([ExcelArgument("Enter true for a row list")]bool isCol = true)
        {
            string[] spNames;
            try
            { spNames = Therm.ListSpecies(); }
            catch (Exception e)
            { return myExcelInt.returnExceptionMessage("Exception in ListSpecies  --  ", e); }

            object[,] ret;
            if (isCol)
                ret = myExcelInt.oneDToColArray(spNames);
            else
                ret = myExcelInt.oneDToRowArray(spNames);

            return ret;
        }

        [ExcelFunction(Description = "returns log10 of the equilibrium constant of formation")]
        public static object logKf_([ExcelArgument("temperature, C")]double t, 
                                [ExcelArgument("specie name")]string spName)
        {
            object otdr = ML.CallHandler.ThermoJANAF(spName, t);
            if (otdr.GetType() != typeof(Thermo.ThermoDataRecord))
            {
                object[] retS = new object[1];
                retS[0] = otdr;
                return myExcelInt.retRowArray(retS);
            }

            Thermo.ThermoDataRecord tdr = (Thermo.ThermoDataRecord)otdr;
            return tdr.log10Kf;
        }

        [ExcelFunction(Description = "returns the enthalpy")]
        public static object Enthalpy_([ExcelArgument("temperature, C")]double tC,
                        [ExcelArgument("specie name")]string spName)
        {
          return ML.CallHandler.Enthalpy(tC, spName);
        }


        [ExcelFunction(Description = "Calculates chemical equilibrium for a given temperature and pressure, kmol/kg")]
        public static object EquilibriumTP_(  [ExcelArgument("temperature, C")]double t, 
                                            [ExcelArgument("pressure, bar gage")]double p,
                                            [ExcelArgument("atoms H")]double aH,
                                            [ExcelArgument("atoms C")]double aC,
                                            [ExcelArgument("atoms N")]double aN,
                                            [ExcelArgument("atoms O")]double aO,
                                            [ExcelArgument("atoms Ar")]double aAr,
                                            [ExcelArgument("atoms s")]double aS)
        {
            double[] myAtoms = new double[6];
            myAtoms[0] = aH;
            myAtoms[1] = aC;
            myAtoms[2] = aN;
            myAtoms[3] = aO;
            myAtoms[4] = aAr;
            myAtoms[5] = aS;
            object[] myOut;
            myOut = ML.CallHandler.EquilibriumTP_(t, p, myAtoms);
            return myExcelInt.retColArray(myOut);
        }

        [ExcelFunction(Description = "Calculates chemical equilibrium for a given temperature and pressure, kmol/kg")]
        public static object EquilibriumTPR_([ExcelArgument("temperature, C")]double t,
                                           [ExcelArgument("pressure, bar gage")]double p,
                                           [ExcelArgument("atoms H")]double aH,
                                           [ExcelArgument("atoms C")]double aC,
                                           [ExcelArgument("atoms N")]double aN,
                                           [ExcelArgument("atoms O")]double aO,
                                           [ExcelArgument("atoms Ar")]double aAr,
                                           [ExcelArgument("atoms s")]double aS)
        {
            double[] myAtoms = new double[6];
            myAtoms[0] = aH;
            myAtoms[1] = aC;
            myAtoms[2] = aN;
            myAtoms[3] = aO;
            myAtoms[4] = aAr;
            myAtoms[5] = aS;
            object[] myOut;
            myOut = ML.CallHandler.EquilibriumTPR_(t, p, myAtoms);
            return myExcelInt.retColArray(myOut);
        }

        [ExcelFunction(Description = "Thermo data consistent with JANAF tables")]
        public static object[,] Thermo_JANAF([ExcelArgument("temperature, C")] double t, [ExcelArgument("specie name")] string spName )
        {
            object otdr = ML.CallHandler.ThermoJANAF(spName, t);
            if (otdr.GetType() != typeof(Thermo.ThermoDataRecord))
            {
                object[] retS = new object[1];
                retS[0] = otdr;
                return myExcelInt.retRowArray(retS);
            }

            Thermo.ThermoDataRecord tdr = (Thermo.ThermoDataRecord)otdr;
            object[] ret = new object[7];
            ret[0] = tdr.cP;
            ret[1] = tdr.s;
            ret[2] = tdr.gHTRef_T;
            ret[3] = tdr.hHTRef / 1000;
            ret[4] = tdr.deltaHf / 1000;
            ret[5] = tdr.deltaGf / 1000;
            ret[6] = tdr.log10Kf;

            return myExcelInt.retRowArray(ret);
        }

        [ExcelFunction(Description = "Thermo data consistent with CEA Thermo Build output")]
        public static object[,] Thermo_CEA([ExcelArgument("temperature, C")] double t, [ExcelArgument("specie name")] string spName)
        {
            object otdr = ML.CallHandler.ThermoJANAF(spName, t);
            if (otdr.GetType() != typeof(Thermo.ThermoDataRecord))
            {
                object[] retS = new object[1];
                retS[0] = otdr;
                return myExcelInt.retRowArray(retS);
            }

            Thermo.ThermoDataRecord tdr = (Thermo.ThermoDataRecord)otdr;
            object[] ret = new object[7];
            int spIndex = Thermo.findSpecieIndex(spName);
            ret[0] = tdr.cP;
            ret[1] = tdr.hHTRef / 1000;
            ret[2] = tdr.s;
            ret[3] = tdr.gHTRef_T;
            ret[4] = Thermo.allSpecies[spIndex].heatForm/1000 + tdr.hHTRef / 1000;
            ret[5] = tdr.deltaHf / 1000;
            ret[6] = tdr.log10Kf;

            return myExcelInt.retRowArray(ret);
        }

        public static object MolWt_(string name)
        {
            string[] spName = new string[] { name };
            return ML.CallHandler.MolWt(spName);
        }
    }
}
