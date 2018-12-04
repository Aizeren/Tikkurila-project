namespace SimplifiedCore
{
    public static class Strings
    {
        public const string UnmanagedDLLUsageFunctionsLibrary = "kernel32.dll";

        public static class ConfigCurrentDirectory
        {
            public const string ConfigFile = "configfile";
            public const string Executable = "executable";
            public const string Path = "dir";
        }

        public const string ManagedDllClassName = "SimplifiedCoreExternalInterface.Methods";

        public static class DLLMethodsNames
        {
            public const string AcceptConnection = "AcceptConnection";
            public const string IsReceiver = "IsReceiver";
            public const string GetMID = "GetMID";
            public const string MatchMID = "MatchMID";
            public const string GetData = "GetData";
            public const string DispatchData = "DispatchData";
            public const string StopAccepting = "StopAccepting";
            public const string Terminate = "Terminate";
            public const string CloseConnection = "CloseConnection";
        }
    }

    public static class Numerics
    {
        public static class BufferSizes
        {
            public const int DataTransferBuffer = 1024;
            public const int GetMIDStringBuilderBuffer = 1024;
        }

        // all the timeouts are expected to be received in milliseconds
        public static class Timeouts
        {
            public const int DataTransferIterationTimeout = 1;
            public const int OpenTransferWaitForTransferTermination = 5000;
        }
    }

    public static class ErrorCodes
    {
        public const int ERROR_SUCCESS = 0;

        // say, when we call Stop() and didn't call Start() before
        public const int NOT_RUNNING = 0x0010;
        public const int ALREADY_RUNNING = 0x0011;
        public const int PREVIOUS_RUN_DOESNT_TERMINATE = 0x0012;
        public const int PENDING_TERMINATION = 0x0013;
        public const int TERMINATION_NOT_REQUESTED = 0x0014;

        public const int UNMANAGED_DLL = 0x0015;
        public const int MANAGED_DLL = 0x0016;

        public const int DLL_NOT_LOADED = 0x0017;
        public const int DLL_ALREADY_LOADED = 0x0018;
        public const int DLL_COULDNT_BE_LOADED = 0x0019;

        // error codes from 0x0100 to 0x01FF
        // are used for file errors
        public const int FILE_NOT_FOUND  = 0x0100;
        public const int FILE_COULDNT_BE_OPEN = 0x0101; // Maybe expand this to ACCESS_DENIED, ALREADY_IN_USE,..
        public const int DLL_DAMAGED = 0x0102;
        public const int PATH_NOT_FOUND = 0x0103;
    }
}
