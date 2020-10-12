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
        const string notFound = "N/A";
        const string parsingError = "Parsing Error";

        [ExcelFunction(Description = "Find Latin characters within the specified text. Returns the first Latin character or word found or " + notFound + " if no match is found")]
        public static string FindFirstLatin([ExcelArgument(Name = "Text", Description = "Relevant Text")] string txt)
        {
            string pattern = @"[\p{IsBasicLatin} -[\s\d\p{P}]]+"; //Basic Latin EXCLUDING Punctuation, numbers and spaces
            Regex rgx = new Regex(pattern, RegexOptions.None);
            Match m;
            try
            {
                m = rgx.Match(txt);
            }
            catch
            {
                return parsingError;
            }

            if (!m.Success)
                return notFound;
            else
                return m.Value;
        }

        [ExcelFunction(Description = "Find Greek characters within the specified text. Returns the first Greek character or word found or " + notFound + " if no match is found")]
        public static string FindFirstGreek([ExcelArgument(Name = "Text", Description = "Relevant Text")] string txt)
        {
            string pattern = @"[\p{IsGreek}]+";  //Greek does not need the negative classes
            Regex rgx = new Regex(pattern, RegexOptions.None);
            Match m;
            try
            {
                m = rgx.Match(txt);
            }
            catch
            {
                return parsingError;
            }

            if (!m.Success)
                return notFound;
            else
                return m.Value;
        }

        [ExcelFunction(Description = "Find the first occurrence of the Regular expression within the specified text. Returns the first match or " + notFound + " if no match is found")]
        public static string FindFirstRegex(
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
                return parsingError;
            }

            if (!m.Success)
                return notFound;
            else
                return m.Value;
        }

        [ExcelFunction(Description = "Find all Latin characters or words within the specified text. Returns a comma separated list of matches or " + notFound + " if no match is found")]
        public static string FindAllLatin([ExcelArgument(Name = "Text", Description = "Relevant Text")] string txt)
        {
            string fullMatch = "";
            string pattern = @"(?<elLatino>[\p{IsBasicLatin} -[\s\d\p{P}]]+)";
            Regex rgx = new Regex(pattern, RegexOptions.ExplicitCapture);
            Match m;
            try
            {
                m = rgx.Match(txt);
            }
            catch
            {
                return parsingError;
            }

            if (!m.Success)
                return notFound;

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


        [ExcelFunction(Description = "Find all Greek characters or words within the specified text. Returns a comma separated list of matches or " + notFound + " if no match is found")]
        public static string FindAllGreek([ExcelArgument(Name = "Text", Description = "Relevant Text")] string txt)
        {
            string fullMatch = "";
            string pattern = @"(?<elGreco>[\p{IsGreek}]+)"; //Greek does not need the negative classes
            Regex rgx = new Regex(pattern, RegexOptions.ExplicitCapture);
            Match m;
            try
            {
                m = rgx.Match(txt);
            }
            catch
            {
                return parsingError;
            }

            if (!m.Success)
                return notFound;

            while (m.Success)
            {
                //Groups[0] is the full match and 1+ contains the named groups
                for (int i = 1; i < m.Groups.Count; i++)
                {
                    fullMatch += m.Groups[i].Value + ",";
                }
                m = m.NextMatch();
            }

            return fullMatch.EndsWith(",") ? fullMatch.Remove(fullMatch.Length - 1) : fullMatch;
        }

        [ExcelFunction(Description = "Find all occurrences of the Regular expression within the specified text. Returns the first match or " + notFound + " if no match is found")]
        public static string FindAllRegex(
                [ExcelArgument(Name = "Text", Description = "Relevant Text")] string txt,
                [ExcelArgument(Name = "Pattern", Description = "Regex Pattern (.NET Compatible)")] string pattern
        )
        {
            string fullMatch = "";
            string newPattern = @"(?<elRegExo>" + pattern + ")";
            Regex rgx = new Regex(newPattern, RegexOptions.ExplicitCapture);
            Match m;
            try
            {
                m = rgx.Match(txt);
            }
            catch
            {
                return parsingError;
            }

            if (!m.Success)
                return notFound;

            while (m.Success)
            {
                //Groups[0] is the full match and 1+ contains the named groups
                for (int i = 1; i < m.Groups.Count; i++)
                {
                    fullMatch += m.Groups[i].Value + ",";
                }
                m = m.NextMatch();
            }

            return fullMatch.EndsWith(",") ? fullMatch.Remove(fullMatch.Length - 1) : fullMatch;
        }
    }
}
