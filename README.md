# Excel-Cell-Checker
A few quick User Defined functions to bring the power of Regular Expressions to detect rogue text and characters in Excel.
The primary goal of the project was to help detect the presence of Greek characters in Latin text and vice versa in an easy and user friendly way

We begin this project with a few simple functions:
1. FindFirstLatin - The first match of a group of Latin characters (words) found within the selected (assumed Greek) text.
2. FindFirstGreek - The first match of a group of Greek characters (words) found within the selected (assumed Latin) text.
3. FindFirstRegex - The first match of the supplied .NET compatible Regular Expression pattern within the slected text.
4. FindAllLatin - Returns a comma separated list of Latin characters (words) within the selected text.
5. FindAllGreek - Returns a comma separated list of Greek characters (words) within the selected text.
6. FindAllRegex - Returns a comma separated list of all matches of the supplied .NET compatible Regular Expression pattern within the slected text.

## Note on Regular Expressions
The FindFirstRegex and FindAllRegex functions are meant to be used with <a href="https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expressions" target="_blank">.NET compatible regular expressions</a> including Exclusion groups! Special features of other Regex parsers will not work and might generate errors.
