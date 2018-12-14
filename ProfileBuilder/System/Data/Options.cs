using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ProfileBuilder.System.Data
{
    class Options
    {
        //[Option('r', Required = true, HelpText = "Input files to be processed.")]
        //public string InputFiles { get; set; }

        // Omitting long name, defaults to name of property, ie "--verbose"
        [Option(
          Default = false,
          HelpText = "Prints all messages to standard output.")]
        public bool Verbose { get; set; }

        [Option("stdin",
          Default = false,
      
          HelpText = "Read from stdin")]
        public bool stdin { get; set; }

        [Value(0, MetaName = "offset", HelpText = "File offset.")]
        public long? Offset { get; set; }
    }
}
