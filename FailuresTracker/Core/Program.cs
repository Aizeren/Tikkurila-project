using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

using System.Reflection;

namespace Core
{
    class Program
    {
        struct StartupParams
        {
            // whether we read DLLs data from config.cfg or another file or from console
            // CONSOLE
            // THIS FILE
            // ANOTHER FILE <CR;LF> <Path>
            public enum EnumDllPathsSource
            {
                Console,
                ThisFile,
                AnotherFile
            }

            public EnumDllPathsSource DllPathsSource;

            public string DllPathsFile;

            // whether start core right away or start pending for manual startup
            // START IMMEDIATELY
            // START MANUALLY
            public bool StartImmediately;

            // gets filled in ReadConfigFile only if
            // DllPathsSource == ThisFile
            public List<string> DllPaths;
        }

        static StartupParams ReadConfigFile( string configFilePath )
        {
            StartupParams retval = new StartupParams();

            // read file string by string
            string[] fileDataRawText = File.ReadAllLines( configFilePath );

            List<string> fileData = new List<string>();

            // remove empty and whitespaces strings
            for (int i = 0; i < fileDataRawText.Length; i++)
            {
                // check if the string does not consist of whitespaces only
                string candidateString = fileDataRawText[i].Trim();

                // in the case when this condition
                // is satisfied, the string is neither
                // empty nor consists of whitespaces only
                if (candidateString.Length != 0)
                {
                    fileData.Add(fileDataRawText[i]);
                }
            }

            switch (fileData[0])
            {
                case "CONSOLE":
                    retval.DllPathsSource = StartupParams.EnumDllPathsSource.Console;
                    break;
                case "THIS FILE":
                    retval.DllPathsSource = StartupParams.EnumDllPathsSource.ThisFile;
                    break;
                case "ANOTHER FILE":
                    retval.DllPathsSource = StartupParams.EnumDllPathsSource.AnotherFile;
                    break;
                default:
                    throw new Exception("Wrong file format - wrong DLLs source option");
            }

            int startupOptionOffset = 1;
            if (retval.DllPathsSource == StartupParams.EnumDllPathsSource.AnotherFile)
            {
                startupOptionOffset++;
                retval.DllPathsFile = fileData[1];
            }

            switch (fileData[startupOptionOffset])
            {
                case "START IMMEDIATELY":
                    retval.StartImmediately = true;
                    break;
                case "START MANUALLY":
                    retval.StartImmediately = false;
                    break;
                default:
                    throw new Exception("Wrong file format - wrong startup option");
            }

            if (retval.DllPathsSource == StartupParams.EnumDllPathsSource.ThisFile)
            {
                for (int i = startupOptionOffset + 1; i < fileData.Count; i++)
                {
                    retval.DllPaths.Add( fileData[i] );
                }
            }

            return retval;
        }


        // Reads DLLs list from console and stores them in
        // ~config.cfg file
        static void ReadFromConsole()
        {
            string current;

            StreamWriter configFile = new StreamWriter("~config.cfg", false, Encoding.ASCII);

            Console.WriteLine("Please, enter paths to the needed DLLs. To finish,\r\npress enter leaving the last line blank.");

            current = Console.ReadLine().Trim();

            bool sustainLoop = true;

            while (sustainLoop)
            {
                while (current != "")
                {
                    configFile.WriteLine(current);
                    current = Console.ReadLine().Trim();
                }

                Console.WriteLine( "Have you finished? (y/n) " );

                char finishing = Console.ReadKey().KeyChar;

                while (finishing != 'y' && finishing != 'n')
                {
                    Console.WriteLine("(y/n) ");
                    finishing = Console.ReadKey().KeyChar;
                }

                sustainLoop = (finishing == 'n');
            }

            configFile.Close();
        }

        static void ReadFromCurrentFile(StartupParams startup)
        {
            StreamWriter configFile = new StreamWriter("~config.cfg", false, Encoding.ASCII);

            foreach (string dllPath in startup.DllPaths)
            {
                configFile.WriteLine(dllPath);
            }

            configFile.Close();
        }
        
        static int Main(string[] args)
        {
            const string initialConfigFilePath = "config.cfg";
            string dllsConfigFilePath = "~config.cfg";

            Core core = new Core();
            
            Console.WriteLine( "Core has just been initialized\r\nReading {0}...", initialConfigFilePath );
            
            StartupParams startup = ReadConfigFile(initialConfigFilePath);

            switch (startup.DllPathsSource)
            {
                case StartupParams.EnumDllPathsSource.Console:
                    ReadFromConsole();
                    break;
                case StartupParams.EnumDllPathsSource.ThisFile:
                    ReadFromCurrentFile(startup);
                    break;
                case StartupParams.EnumDllPathsSource.AnotherFile:
                    dllsConfigFilePath = startup.DllPathsFile;
                    break;
            }

            bool isCoreRunning = startup.StartImmediately;

            if (startup.StartImmediately)
            {
                core.Start(dllsConfigFilePath);
            }

            string command = Console.ReadLine().Trim().ToLower();

            bool sustainLoop = true;
            
            while (sustainLoop)
            {
                switch (command)
                {
                    case "start":
                        isCoreRunning = true;
                        if (core.Start(dllsConfigFilePath) == ErrorCodes.ALREADY_RUNNING)
                        {
                            Console.WriteLine("Already running");
                        }
                        break;
                    case "stop":
                        isCoreRunning = false;
                        if (core.Stop(true) == ErrorCodes.NOT_RUNNING)
                        {
                            Console.WriteLine("Not running");
                        }
                        break;
                    case "exit":
                        if (!isCoreRunning)
                        {
                            sustainLoop = false;
                            continue;
                        }
                        else
                        {
                            Console.WriteLine( "Please, stop the system first" );
                        }
                        break;
                    default:
                        throw new Exception( "Invocation of an unknown command" );
                }
                command = Console.ReadLine().Trim().ToLower();
            }

            return ErrorCodes.ERROR_SUCCESS;
        }
    }
}