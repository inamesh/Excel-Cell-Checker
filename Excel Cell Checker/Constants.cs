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
    public static class Constants
    {
        public const string NotFound = "N/A";
        public const string ParsingError = "Parsing Error";
        public const string LatinPattern = @"[\p{IsBasicLatin} -[\s\d\p{P}]]+";
        public const string GreekPattern = @"[\p{IsGreek}]+";
    }
}
