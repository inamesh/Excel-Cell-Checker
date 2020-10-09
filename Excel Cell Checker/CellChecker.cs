using System.Text.RegularExpressions;
using ExcelDna.Integration;

/*
 * 
 * 
 * NB: The patterns here use character class subtraction in .NET specific format and will not work with any other parser 
 * Documentation: https://docs.microsoft.com/en-us/dotnet/standard/base-types/character-classes-in-regular-expressions#character-class-subtraction-base_group---excluded_group
 * 
 * 
 * 
 */

namespace Excel_Cell_Checker
{
    public static class CellChecker
    {
        [ExcelFunction(Description = "Find the first Latin character within the specified text")]
        public static string CheckForLatin([ExcelArgument(Name = "Text", Description = "Relevant Text")] string txt)
        {
            string pattern = @"[\p{IsBasicLatin} -[\s\d\p{P}]]+"; 
            Regex rgx = new Regex(pattern, RegexOptions.None);
            Match m;
            try
            {
                m = rgx.Match(txt);
            }
            catch
            {
                return "Parsing Error";
            }
            if (!m.Success)
            {
                return "No match found!";
            }
            return m.Value;
        }

        [ExcelFunction(Description = "Find the first Greek character within the specified text")]
        public static string CheckForGreek([ExcelArgument(Name = "Text", Description = "Relevant Text")] string txt)
        {
            string pattern = @"[\p{IsGreek} -[\s\d\p{P}]]+";
            Regex rgx = new Regex(pattern, RegexOptions.None);
            Match m;
            try
            {
                m = rgx.Match(txt);
            }
            catch
            {
                return "Parsing Error";
            }
            if (!m.Success)
            {
                return "No match found!";
            }
            return m.Value;
        }

        [ExcelFunction(Description = "Find the first occurrence of the Regular expression within the specified text")]
        public static string CheckRegex(
            [ExcelArgument(Name = "Text", Description = "Relevant Text")] string txt,
            [ExcelArgument(Name = "Pattern", Description = "Regex Pattern (.NET Compatible)")] string pattern
        )
        {
            Regex rgx = new Regex(pattern, RegexOptions.None);
            Match m;
            try
            {
                m = rgx.Match(txt);
            }
            catch
            {
                return "Parsing Error";
            }
            if (!m.Success)
            {
                return "No match found!";
            }
            return m.Value;
        }
    }
}
