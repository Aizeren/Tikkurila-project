using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Reflection;

namespace Core
{
    class ManagedDLLLoader : IDLLLoader
    {
        /*
         * We want to properly unload
         * DLLs when closing an interface,
         * so every IDLLLoader must do this
         * in UnloadDLL function.
         * 
         * But with managed DLLs, that are
         * loaded as assemblies, it's a little
         * tricky: an assembly cannot be unloaded
         * manually - it gets removed automatically
         * when the AppDomain they were loaded to,
         * gets unloaded.
         * 
         * So, we create an app domain, so to say,
         * "local" for this particular DLLLoader.
         * In this case, when UnloadDLL is called,
         * we can just use AppDomain.Unload( _LocalDomain )
         */
        //private AppDomain _LocalDomain;
        private Assembly _DllAssembly;

        // The class, specified in the dll,
        // that contains needed functions
        private Type _DllMainClass;

        public ManagedDLLLoader()
        {
            _DllAssembly = null;
        }

        public int LoadDLL(string dllPath)
        {
            if (_DllAssembly != null)
            {
                return ErrorCodes.DLL_ALREADY_LOADED;
            }

            if (!File.Exists(dllPath))
            {
                return ErrorCodes.FILE_NOT_FOUND;
            }

            //System.Security.Policy.Evidence evd = new System.Security.Policy.Evidence();
            /*
            _LocalDomain = AppDomain.CreateDomain(
                dllPath,
                evd,
                Path.GetDirectoryName( dllPath ),
                "",
                false
                );
                */
            _DllAssembly =
                Assembly.LoadFile( dllPath );

            if (_DllAssembly == null)
            {
                return ErrorCodes.DLL_COULDNT_BE_LOADED;
            }

            _DllMainClass = _DllAssembly.GetType( Strings.ManagedDllClassName );

            return ErrorCodes.ERROR_SUCCESS;
        }

        public AcceptConnection_Delegate GetAcceptConnectionDelegate()
        {
            return
                (AcceptConnection_Delegate)
                _DllMainClass.GetMethod(Strings.DLLMethodsNames.AcceptConnection)
                .CreateDelegate(typeof(AcceptConnection_Delegate));
        }

        public CloseConnection_Delegate GetCloseConnectionDelegate()
        {
            return
                (CloseConnection_Delegate)
                _DllMainClass.GetMethod(Strings.DLLMethodsNames.CloseConnection)
                .CreateDelegate(typeof(CloseConnection_Delegate));
        }

        public DispatchData_Delegate GetDispatchDataDelegate()
        {
            return
                (DispatchData_Delegate)
                _DllMainClass.GetMethod(Strings.DLLMethodsNames.DispatchData)
                .CreateDelegate(typeof(DispatchData_Delegate));
        }

        public GetData_Delegate GetGetDataDelegate()
        {
            return
                (GetData_Delegate)
                _DllMainClass.GetMethod(Strings.DLLMethodsNames.GetData)
                .CreateDelegate(typeof(GetData_Delegate));
        }

        public GetMID_Delegate GetGetMIDDelegate()
        {
            return
                (GetMID_Delegate)
                _DllMainClass.GetMethod(Strings.DLLMethodsNames.GetMID)
                .CreateDelegate(typeof(GetMID_Delegate));
        }

        public IsReceiver_Delegate GetIsReceiverDelegate()
        {
            return
                (IsReceiver_Delegate)
                _DllMainClass.GetMethod(Strings.DLLMethodsNames.IsReceiver)
                .CreateDelegate(typeof(IsReceiver_Delegate));
        }

        public MatchMID_Delegate GetMatchMIDDelegate()
        {
            return
                (MatchMID_Delegate)
                _DllMainClass.GetMethod(Strings.DLLMethodsNames.MatchMID)
                .CreateDelegate(typeof(MatchMID_Delegate));
        }

        public StopAccepting_Delegate GetStopAcceptingDelegate()
        {
            return
                (StopAccepting_Delegate)
                _DllMainClass.GetMethod(Strings.DLLMethodsNames.StopAccepting)
                .CreateDelegate(typeof(StopAccepting_Delegate));
        }

        public Terminate_Delegate GetTerminateDelegate()
        {
            return
                (Terminate_Delegate)
                _DllMainClass.GetMethod(Strings.DLLMethodsNames.Terminate)
                .CreateDelegate(typeof(Terminate_Delegate));
        }

        public int UnloadDLL()
        {
            if (_DllAssembly == null)
            {
                return ErrorCodes.DLL_NOT_LOADED;
            }

            _DllAssembly = null;

            return ErrorCodes.ERROR_SUCCESS;
        }
    }
}
