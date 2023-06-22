using CsvHelper;
using Fclp;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace wmi
{
    /// <summary>
    /// 
    /// </summary>
    internal class ApplicationArguments
    {
        public string Output { get; set; }
        public string Input { get; set; }
    }

    class Program
    {
        #region Member Variables
        private static FluentCommandLineParser<ApplicationArguments> fclp;
        #endregion

        /// <summary>
        /// Application entry point
        /// </summary>
        /// <param name="args"></param>
        // Define constants
        const string EventConsumerPattern = @"([\w_]*EventConsumer)\.Name=""([\w\s]*)""";
        const string EventFilterPattern = @"_EventFilter\.Name=""([\w\s]*)""";
        const string CommandLineEventConsumerPattern = @"\x00CommandLineEventConsumer\x00\x00(.*?)\x00.*?{0}\x00\x00?([^\x00]*)?";
        const string EventConsumerPatternWithGroupName = @"(\w*EventConsumer)(.*?)({0})(\x00\x00)([^\x00]*)(\x00\x00)([^\x00]*)";
        const string EventConsumerPatternWithFilter = @"({0})(\x00\x00)([^\x00]*)(\x00\x00)";

        static void Main(string[] args)
        {
            if (!ProcessCommandLine(args))
            {
                return;
            }

            if (!CheckCommandLine())
            {
                return;
            }

            var fileInfo = new FileInfo(fclp.Object.Input);
            Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fffffff} | Parsing file: {fileInfo.Name}");
            Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fffffff} | File size: {fileInfo.Length / 1024.0 / 1024.0:F2} MB");

            string data = File.ReadAllText(fileInfo.FullName);

            Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fffffff} | Searching for pattern: {EventConsumerPattern}");
            Regex regexConsumer = new Regex(EventConsumerPattern, RegexOptions.Multiline | RegexOptions.IgnoreCase | RegexOptions.Compiled);
            var matchesConsumer = regexConsumer.Matches(data);
            Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fffffff} | Found {matchesConsumer.Count} matches for pattern: {EventConsumerPattern}");

            Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fffffff} | Searching for pattern: {EventFilterPattern}");
            Regex regexFilter = new Regex(EventFilterPattern, RegexOptions.Multiline | RegexOptions.IgnoreCase | RegexOptions.Compiled);
            var matchesFilter = regexFilter.Matches(data);
            Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fffffff} | Found {matchesFilter.Count} matches for pattern: {EventFilterPattern}");

            List<Binding> bindings = new List<Binding>();
            for (int index = 0; index < matchesConsumer.Count; index++)
            {
                bindings.Add(new Binding(matchesConsumer[index].Groups[2].Value, matchesFilter[index].Groups[1].Value));
            }

            foreach (var b in bindings)
            {
                var cmdLineEventConsPattern = string.Format(CommandLineEventConsumerPattern, b.Name);
                Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fffffff} | Searching for pattern: {cmdLineEventConsPattern}");
                Regex regexEventConsumer = new Regex(cmdLineEventConsPattern, RegexOptions.Multiline);
                var matches = regexEventConsumer.Matches(data);
                Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fffffff} | Found {matches.Count} matches for pattern: {cmdLineEventConsPattern}");

                foreach (Match m in matches)
                {
                    b.Type = "CommandLineEventConsumer";
                    b.Arguments = m.Groups[1].Value;
                }

                var eventConsGroupPattern = string.Format(EventConsumerPatternWithGroupName, b.Name);
                Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fffffff} | Searching for pattern: {eventConsGroupPattern}");
                regexEventConsumer = new Regex(eventConsGroupPattern, RegexOptions.Multiline);
                matches = regexEventConsumer.Matches(data);
                Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fffffff} | Found {matches.Count} matches for pattern: {eventConsGroupPattern}");

                foreach (Match m in matches)
                {
                    b.Other = $"{m.Groups[1]} ~ {m.Groups[3]} ~ {m.Groups[5]} ~ {m.Groups[7]}";
                }

                var eventConsFilterPattern = string.Format(EventConsumerPatternWithFilter, b.Filter);
                Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fffffff} | Searching for pattern: {eventConsFilterPattern}");
                regexEventConsumer = new Regex(eventConsFilterPattern, RegexOptions.Multiline);
                matches = regexEventConsumer.Matches(data);
                Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fffffff} | Found {matches.Count} matches for pattern: {eventConsFilterPattern}");

                foreach (Match m in matches)
                {
                    b.Query = m.Groups[3].Value;
                }
            }

            OutputToConsole(bindings);

            if (fclp.Object.Output != null)
            {
                OutputToFile(bindings);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private static bool ProcessCommandLine(string[] args)
        {
            fclp = new FluentCommandLineParser<ApplicationArguments>
            {
                IsCaseSensitive = false
            };

            fclp.Setup(arg => arg.Input)
               .As('i')
               .Required()
               .WithDescription("Input file (OBJECTS.DATA)");

            fclp.Setup(arg => arg.Output)
                .As('o')
                .WithDescription("Output directory for analysis results");

            var header =
               $"\r\n{Assembly.GetExecutingAssembly().GetName().Name} v{Assembly.GetExecutingAssembly().GetName().Version.ToString(3)}" +
               "\r\n\r\nAuthor: Mark Woan / woanware (markwoan@gmail.com)" +
               "\r\nhttps://github.com/woanware/wmi-parser" +
               "\r\n\r\nUpdated By: Andrew Rathbun / https://github.com/AndrewRathbun" +
               "\r\nhttps://github.com/AndrewRathbun/wmi-parser";

            // Sets up the parser to execute the callback when -? or --help is supplied
            fclp.SetupHelp("?", "help")
                .WithHeader(header)
                .Callback(text => Console.WriteLine(text));

            var result = fclp.Parse(args);

            if (result.HelpCalled)
            {
                return false;
            }

            if (result.HasErrors)
            {
                Console.WriteLine("");
                Console.WriteLine(result.ErrorText);
                fclp.HelpOption.ShowHelp(fclp.Options);
                return false;
            }

            Console.WriteLine(header);
            Console.WriteLine("");

            return true;
        }

        /// <summary>
        /// Performs some basic command line parameter checking
        /// </summary>
        /// <returns></returns>
        private static bool CheckCommandLine()
        {
            if (Directory.Exists(fclp.Object.Input))
            {
                var dirInfo = new DirectoryInfo(fclp.Object.Input);
                if (!dirInfo.EnumerateFiles().Any())
                {
                    Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fffffff} | Input directory (-i) is empty.");
                    return false;
                }
            }
            else if (File.Exists(fclp.Object.Input))
            {
                // nothing to do, file exists
            }
            else
            {
                Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fffffff} | Input (-i) does not exist.");
                return false;
            }

            if (fclp.Object.Output != null)
            {
                if (!Directory.Exists(fclp.Object.Output))
                {
                    Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fffffff} | Output directory (-o) does not exist, creating it...");
                    Directory.CreateDirectory(fclp.Object.Output);
                }
            }

            return true;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="bindings"></param>
        private static void OutputToConsole(List<Binding> bindings)
        {
            // Output the data
            foreach (var b in bindings)
            {
                if ((b.Name.Contains("BVTConsumer") && b.Filter.Contains("BVTFilter")) || (b.Name.Contains("SCM Event Log Consumer") && b.Filter.Contains("SCM Event Log Filter")))
                {
                    Console.WriteLine("  {0}-{1} - (Common binding based on consumer and filter names, possibly legitimate)", b.Name, b.Filter);
                }
                else
                {
                    Console.WriteLine("  {0}-{1}\n", b.Name, b.Filter);
                }

                if (b.Type == "CommandLineEventConsumer")
                {
                    Console.WriteLine("    Name: {0}", b.Name);
                    Console.WriteLine("    Type: {0}", "CommandLineEventConsumer");
                    Console.WriteLine("    Arguments: {0}", b.Arguments);
                }
                else
                {
                    Console.WriteLine("    Consumer: {0}", b.Other);
                }

                Console.WriteLine("\n    Filter:");
                Console.WriteLine("      Filter Name : {0}     ", b.Filter);
                Console.WriteLine("      Filter Query: {0}     ", b.Query);
                Console.WriteLine("");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="bindings"></param>
        private static void OutputToFile(List<Binding> bindings)
        {
            using (var writer = new StreamWriter(Path.Combine(fclp.Object.Output, "wmi-parser.tsv")))
            {
                var config = new CsvHelper.Configuration.CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    Delimiter = "\t",
                };
                using (var csv = new CsvWriter(writer, config))
                {
                    // Write out the file headers
                    csv.WriteField("Name");
                    csv.WriteField("Type");
                    csv.WriteField("Arguments");
                    csv.WriteField("Filter Name");
                    csv.WriteField("Filter Query");
                    csv.NextRecord();

                    foreach (var b in bindings)
                    {
                        csv.WriteField(b.Name);
                        csv.WriteField(b.Type);
                        csv.WriteField(b.Arguments);
                        csv.WriteField(b.Filter);
                        csv.WriteField(b.Query);
                        csv.NextRecord();
                    }
                }
            }
        }
    }
}
