using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimplifiedCore
{
    interface IDLLLoader
    {
        /*
         * This function is used to load a DLL.
         * CONSTRUCTOR JUST MAKES AN OBJECT,
         * IT DOES NOT LOAD A DLL
         */
        int LoadDLL(String dllPath);

        AcceptConnection_Delegate GetAcceptConnectionDelegate();

        IsReceiver_Delegate GetIsReceiverDelegate();

        GetMID_Delegate GetGetMIDDelegate();

        MatchMID_Delegate GetMatchMIDDelegate();

        GetData_Delegate GetGetDataDelegate();

        DispatchData_Delegate GetDispatchDataDelegate();

        StopAccepting_Delegate GetStopAcceptingDelegate();

        Terminate_Delegate GetTerminateDelegate();

        CloseConnection_Delegate GetCloseConnectionDelegate();

        int UnloadDLL();
    }
}
