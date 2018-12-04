using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.InteropServices;

using System.IO;

namespace SimplifiedCore
{
    class UnmanagedDLLLoader : IDLLLoader
    {
        /*
         * The functions that are used to communicate
         * with an external Unmanaged DLL
         */
        [DllImport(Strings.UnmanagedDLLUsageFunctionsLibrary)]
        private static extern IntPtr LoadLibrary(string dllPath);

        [DllImport(Strings.UnmanagedDLLUsageFunctionsLibrary, SetLastError = true)]
        private static extern IntPtr GetProcAddress(IntPtr dllHandle, string procName);

        [DllImport(Strings.UnmanagedDLLUsageFunctionsLibrary, SetLastError = true)]
        private static extern bool FreeLibrary(IntPtr dllHandle);

        private IntPtr _DllHandle;

        private IntPtr _CloseConnectionPtr;

        private IntPtr _AcceptConnectionPtr;
        private IntPtr _StopAcceptingPtr;
        private IntPtr _TerminatePtr;

        private IntPtr _IsReceiverPtr;
        private IntPtr _GetDataPtr;
        private IntPtr _DispatchDataPtr;
        private IntPtr _GetMIDPtr;
        private IntPtr _MatchMIDPtr;

        public UnmanagedDLLLoader()
        {
            _DllHandle = IntPtr.Zero;
        }

        public int LoadDLL(String dllPath)
        {
            if (_DllHandle != IntPtr.Zero)
            {
                return ErrorCodes.DLL_ALREADY_LOADED;
            }

            if (!File.Exists(dllPath))
            {
                return ErrorCodes.FILE_NOT_FOUND;
            }

            _DllHandle = LoadLibrary(dllPath);

            if (_DllHandle == IntPtr.Zero)
            {
                return ErrorCodes.DLL_COULDNT_BE_LOADED;
            }

            // Loading functions from the unmanaged DLL
            _IsReceiverPtr = GetProcAddress(_DllHandle, Strings.DLLMethodsNames.IsReceiver);
            _GetDataPtr = GetProcAddress(_DllHandle, Strings.DLLMethodsNames.GetData);
            _DispatchDataPtr = GetProcAddress(_DllHandle, Strings.DLLMethodsNames.DispatchData);
            _GetMIDPtr = GetProcAddress(_DllHandle, Strings.DLLMethodsNames.GetMID);
            _MatchMIDPtr = GetProcAddress(_DllHandle, Strings.DLLMethodsNames.MatchMID);
            _AcceptConnectionPtr = GetProcAddress(_DllHandle, Strings.DLLMethodsNames.AcceptConnection);
            _StopAcceptingPtr = GetProcAddress( _DllHandle, Strings.DLLMethodsNames.StopAccepting );
            _TerminatePtr = GetProcAddress(_DllHandle, Strings.DLLMethodsNames.Terminate);
            _CloseConnectionPtr = GetProcAddress(_DllHandle, Strings.DLLMethodsNames.CloseConnection);

            return ErrorCodes.ERROR_SUCCESS;
        }

        public AcceptConnection_Delegate GetAcceptConnectionDelegate()
        {
            return (AcceptConnection_Delegate)
                Marshal.GetDelegateForFunctionPointer(_AcceptConnectionPtr, typeof(AcceptConnection_Delegate));
        }

        public IsReceiver_Delegate GetIsReceiverDelegate()
        {
            return (IsReceiver_Delegate)
                Marshal.GetDelegateForFunctionPointer(_IsReceiverPtr, typeof(IsReceiver_Delegate));
        }

        public GetMID_Delegate GetGetMIDDelegate()
        {
            return (GetMID_Delegate)
                Marshal.GetDelegateForFunctionPointer(_GetMIDPtr, typeof(GetMID_Delegate));
        }

        public MatchMID_Delegate GetMatchMIDDelegate()
        {
            return (MatchMID_Delegate)
                Marshal.GetDelegateForFunctionPointer(_MatchMIDPtr, typeof(MatchMID_Delegate));
        }

        public GetData_Delegate GetGetDataDelegate()
        {
            return (GetData_Delegate)
                Marshal.GetDelegateForFunctionPointer(_GetDataPtr, typeof(GetData_Delegate));
        }

        public DispatchData_Delegate GetDispatchDataDelegate()
        {
            return (DispatchData_Delegate)
                Marshal.GetDelegateForFunctionPointer(_DispatchDataPtr, typeof(DispatchData_Delegate));
        }

        public StopAccepting_Delegate GetStopAcceptingDelegate()
        {
            return (StopAccepting_Delegate)
                Marshal.GetDelegateForFunctionPointer( _StopAcceptingPtr, typeof(StopAccepting_Delegate) );
        }

        public Terminate_Delegate GetTerminateDelegate()
        {
            return (Terminate_Delegate)
                Marshal.GetDelegateForFunctionPointer(_TerminatePtr, typeof(Terminate_Delegate));
        }

        public CloseConnection_Delegate GetCloseConnectionDelegate()
        {
            return (CloseConnection_Delegate)
                Marshal.GetDelegateForFunctionPointer(_CloseConnectionPtr, typeof(CloseConnection_Delegate));
        }

        public int UnloadDLL()
        {
            if (_DllHandle == IntPtr.Zero)
            {
                return ErrorCodes.DLL_NOT_LOADED;
            }

            FreeLibrary(_DllHandle);

            return ErrorCodes.ERROR_SUCCESS;
        }
    }
}
