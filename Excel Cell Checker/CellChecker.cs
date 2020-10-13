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

        private static string FindSingle(string text, string pattern)
        {
            Regex rgx = new Regex(pattern, RegexOptions.None);
            Match m;
            try
            {
                m = rgx.Match(text);
            }
            catch
            {
                return Constants.ParsingError;
            }

            if (!m.Success)
                return Constants.NotFound;
            else
                return m.Value;
        }

        private static string FindAll(string text, string pattern)
        {
            string newPattern = @"(?<elRegExo>" + pattern + ")";
            string fullMatch = "";
            Regex rgx = new Regex(newPattern, RegexOptions.ExplicitCapture);
            Match m;
            try
            {
                m = rgx.Match(text);
            }
            catch
            {
                return Constants.ParsingError;
            }

            if (!m.Success)
                return Constants.NotFound;

            while (m.Success)
            {
                //Groups[0] is the full match and 1+ contains the named groups
                for (int i = 1; i < m.Groups.Count; i++)
                {
                    fullMatch += m.Groups[i].Value + ",";
                }
                m = m.NextMatch();
            }

            if (fullMatch.EndsWith(","))
                return fullMatch.Remove(fullMatch.Length - 1);
            else
                return fullMatch;
        }

        [ExcelFunction(Description = "Find Latin characters within the specified text. Returns the first Latin character or word found or " + Constants.NotFound + " if no match is found")]
        public static string FindFirstLatin([ExcelArgument(Name = "Text", Description = "Relevant Text")] string txt)
        {
            return FindSingle(txt, Constants.LatinPattern);
        }

        [ExcelFunction(Description = "Find Greek characters within the specified text. Returns the first Greek character or word found or " + Constants.NotFound + " if no match is found")]
        public static string FindFirstGreek([ExcelArgument(Name = "Text", Description = "Relevant Text")] string txt)
        {
            return FindSingle(txt, Constants.GreekPattern);
        }

        [ExcelFunction(Description = "Find the first occurrence of the Regular expression within the specified text. Returns the first match or " + Constants.NotFound + " if no match is found")]
        public static string FindFirstRegex(
            [ExcelArgument(Name = "Text", Description = "Relevant Text")] string txt,
            [ExcelArgument(Name = "Pattern", Description = "Regex Pattern (.NET Compatible)")] string pattern
        )
        {
            return FindSingle(txt, pattern);
        }

        [ExcelFunction(Description = "Find all Latin characters or words within the specified text. Returns a comma separated list of matches or " + Constants.NotFound + " if no match is found")]
        public static string FindAllLatin([ExcelArgument(Name = "Text", Description = "Relevant Text")] string txt)
        {
            return FindAll(txt, Constants.LatinPattern);
        }


        [ExcelFunction(Description = "Find all Greek characters or words within the specified text. Returns a comma separated list of matches or " + Constants.NotFound + " if no match is found")]
        public static string FindAllGreek([ExcelArgument(Name = "Text", Description = "Relevant Text")] string txt)
        {
            return FindAll(txt, Constants.GreekPattern);
        }

        [ExcelFunction(Description = "Find all occurrences of the Regular expression within the specified text. Returns the first match or " + Constants.NotFound + " if no match is found")]
        public static string FindAllRegex(
                [ExcelArgument(Name = "Text", Description = "Relevant Text")] string txt,
                [ExcelArgument(Name = "Pattern", Description = "Regex Pattern (.NET Compatible)")] string pattern
        )
        {
            return FindAll(txt, pattern);
        }
    }
}
