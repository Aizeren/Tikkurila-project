using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Core
{
    interface IDefinedEntity
    {
        // Function called by either
        // Receiver of Sender to get
        // ID of the connected entity
        // (and again, this ID only
        // has sense to DLL)
        void AcceptID(UInt32 id);

        void AcceptMIDDelegates(GetMID_Delegate getMID, MatchMID_Delegate matchMID);

        void AcceptCloseConnectionDelegate(CloseConnection_Delegate closeConnection);
    }
}
