using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Core
{
    class Sender : IDefinedEntity
    {
        private UInt32 _ID;

        private GetMID_Delegate _GetMIDRoutine;
        private MatchMID_Delegate _MatchMIDRoutine;

        private GetData_Delegate _GetDataRoutine;

        private CloseConnection_Delegate _CloseConnectionRoutine;

        public string GetMID()
        {
            StringBuilder buffer = new StringBuilder(Numerics.BufferSizes.GetMIDStringBuilderBuffer);

            _GetMIDRoutine(_ID, buffer);

            return buffer.ToString();
        }

        public bool MatchMID(string id)
        {
            return (_MatchMIDRoutine(_ID, id) != 0);
        }

        public void AcceptID(UInt32 id)
        {
            _ID = id;
        }

        public void AcceptMIDDelegates(GetMID_Delegate getMID, MatchMID_Delegate matchMID)
        {
            _GetMIDRoutine = getMID;
            _MatchMIDRoutine = matchMID;
        }

        public void AcceptGetDataDelegate(GetData_Delegate getData)
        {
            _GetDataRoutine = getData;
        }

        public void AcceptCloseConnectionDelegate(CloseConnection_Delegate closeConnection)
        {
            _CloseConnectionRoutine = closeConnection;
        }

        public void GetData(byte[] data)
        {
            _GetDataRoutine( _ID, data );
        }

        public void CloseConnection()
        {
            _CloseConnectionRoutine( _ID );
        }
    }
}
