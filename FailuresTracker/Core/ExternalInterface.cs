using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.InteropServices;

using System.IO;

namespace Core
{
    /*
     * ExternalInterface class:
     * 
     * Represents methods for controlling
     * an interface that communicates with
     * external entities (we don't call them
     * "devices", because this would be too
     * narrow for the whole system) by some
     * particular means it supports, say,
     * sockets or USB-connection or filesystem -
     * whatever.
     * 
     * The very interface is supposed to be
     * a Dynamic-Link Library (DLL) that exports
     * a set of functions, essential for us
     * to work.
     * 
     * NEEDTOKNOW: if you are implementing your
     * own .dll, you are gonna make all the functions
     * thread-safe YOURSELF.
     * 
     * _Delegate suffix is for types
     * Delegate suffix is for the delegates to be stored and passed
     * Routine suffix is for the delegates to be invoked
     */

    /*
     * Functions, that a DLL must export
     * to be used by the Core
     * 
     * The IDs passed to these functions are the IDs
     * that the DLL uses to identify entities.
     * So, Core knows absolutely nothing about those
     * numbers - our task is only to save them in
     * corresponding classes and pass inside the DLL
     * when we call the functions.
     * First, we get this ID from AcceptConnection() -
     * its returned value.
     */

    /*
     * First line in those eight commentaries below
     * specifies, how those functions must be declared
     * in the DLL that's gonna be loaded and used.
     * 
     * And the declarations are given not in "C style"
     * or "C++ style", but they are pretty self-describing:
     * 
     * Type Name( Parameters )
     * 
     * uint32 type, as an example, states an unsigned 32-bit
     * integer value.
     * 
     * But one hint really must be given:
     * (type1 -> type2) means, that in the
     * essence, the type is supposed to be
     * type1, but due to some incompatibilities
     * or whatever, type2 serves as an
     * "intermediary".
     */

    /*
     * uint32 AcceptConnection()
     * 
     * The function is used to accept
     * those who connect via this interface.
     * It returns the ID that is described
     * in the last paragraph of the previous
     * commentary
     */
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    public delegate UInt32 AcceptConnection_Delegate();

    /*
     * (bool -> uint32) IsReceiver( uint32 id )
     * 
     * When someone connects to an interface, we
     * don't know if it's a Sender or a Receiver,
     * so we call this function to find out
     */
    // type of this function is actually bool, not int, but dll's bool
    // and C#'s bool seem to be incompatible, so we use UInt32
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    public delegate UInt32 IsReceiver_Delegate( UInt32 id );

    /*
     * void GetMID( uint32 id, char * buffer )
     * 
     * Description of what it is for is given
     * in the description of Core
     * 
     * Managed DLLs should fill matchIDBuffer using .Append() function
     */
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    public delegate void GetMID_Delegate( UInt32 id, StringBuilder matchIDBuffer );

    /*
     * (bool -> uint32) MatchMID( uint32 id, char * matchID )
     * 
     * Look at Core description to find out why it is needed
     */
    // type of this function is actually bool, not int, but dll's bool
    // and C#'s bool seem to be incompatible, so we use UInt32
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    public delegate UInt32 MatchMID_Delegate( UInt32 id, string matchID );

    /*
     * void GetData( uint32 id, byte * buffer )
     * 
     * When we call this function, DLL receives data
     * from the sender, specified by "id" parameter
     * and stores it in "dataBuffer"
     */
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    public delegate void GetData_Delegate( UInt32 id, byte[] dataBuffer );

    /*
     * void DispatchData( uint32 id, byte * data )
     * 
     * Sends data from the second argument to the receiver,
     * whose id is passed in the first argument
     */
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    public delegate void DispatchData_Delegate( UInt32 id, byte[] data );

    /*
     * void StopAccepting()
     * 
     * This function makes DLL quit from AcceptConnection().
     * We cannot interrupt execution of that code ourselves,
     * because the DLL can be unmanaged, and threads running
     * unmanaged code are not Abort()ed
     */
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    public delegate void StopAccepting_Delegate();

    /*
     * void Terminate()
     * 
     * This function is called when we close the interface,
     * so we are not going to use it any longer - on DLL side
     * it is supposed to free all the allocated resources,
     * stop all the threads and all that...
     * After we call this functions, we cannot use the DLL
     * any longer (actually we can, but it's gonna be undefined behavior)
     * 
     * NEEDTOKNOW: Don't call this function unless you've ClosedConnenction
     * for every entity (Sender/Receiver).
     */
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    public delegate void Terminate_Delegate();

    /*
     * void CloseConnection( uint32 id )
     * 
     * Closes a certain connection, by ID
     */
    [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
    public delegate void CloseConnection_Delegate( UInt32 id );

    /*
     * Represents an interface (with a corresponding DLL)
     * that sustains Receivers and Senders functionality
     */
    class ExternalInterface : IDisposable
    {
        private IDLLLoader DLLLoader;

        private AcceptConnection_Delegate AcceptConnectionRoutine;
        private IsReceiver_Delegate IsReceiverDelegate;
        private GetMID_Delegate GetMIDDelegate;
        private MatchMID_Delegate MatchMIDDelegate;
        private GetData_Delegate GetDataDelegate;
        private DispatchData_Delegate DispatchDataDelegate;
        private StopAccepting_Delegate StopAcceptingRoutine;
        private Terminate_Delegate TerminateRoutine;
        private CloseConnection_Delegate CloseConnectionDelegate;

        private volatile bool _IsTerminating;

        private int GetDLLType( String dllPath )
        {
            if (!File.Exists(dllPath))
            {
                return ErrorCodes.FILE_NOT_FOUND;
            }

            Stream fileStream = new FileStream( dllPath, FileMode.Open, FileAccess.Read );
            BinaryReader binaryReader = new BinaryReader( fileStream );

            if (fileStream == null || binaryReader == null)
            {
                return ErrorCodes.FILE_COULDNT_BE_OPEN;
            }

            if (fileStream.Length < 64)
            {
                return ErrorCodes.DLL_DAMAGED;
            }

            fileStream.Position = 0x3C;
            uint peHeaderOffset = binaryReader.ReadUInt32();

            if (peHeaderOffset == 0)
            {
                peHeaderOffset = 0x80;
            }

            if (peHeaderOffset > fileStream.Length - 256)
            {
                return ErrorCodes.DLL_DAMAGED; // TODO: Damaged or Not?..
            }

            fileStream.Position = peHeaderOffset;
            uint peHeaderSignature = binaryReader.ReadUInt32();
            if (peHeaderSignature != 0x00004550)
            {
                return ErrorCodes.DLL_DAMAGED;
            }

            fileStream.Position += 20;

            const ushort PE32 = 0x10b;
            const ushort PE32Plus = 0x20b;

            ushort peFormat = binaryReader.ReadUInt16();
            if (peFormat != PE32 && peFormat != PE32Plus)
            {
                return ErrorCodes.DLL_DAMAGED;
            }

            ushort dataDictionaryOffset = (ushort)(peHeaderOffset + (peFormat == PE32 ? 232 : 248));
            fileStream.Position = dataDictionaryOffset;

            uint cliHeaderRva = binaryReader.ReadUInt32();
            if (cliHeaderRva != 0)
            {
                return ErrorCodes.MANAGED_DLL;
            }
            else
            {
                return ErrorCodes.UNMANAGED_DLL;
            }
        }

        public ExternalInterface(String dllPath)
        {
            int statusGetDLLType = GetDLLType( dllPath );

            switch( statusGetDLLType )
            {
                case ErrorCodes.UNMANAGED_DLL:
                    DLLLoader = new UnmanagedDLLLoader();
                    break;
                case ErrorCodes.MANAGED_DLL:
                    DLLLoader = new ManagedDLLLoader();
                    break;
                case ErrorCodes.DLL_DAMAGED:
                    throw new Exception( dllPath + " DLL file is damaged (maybe it's not even a DLL)" );
                case ErrorCodes.FILE_NOT_FOUND:
                    throw new Exception( dllPath + " DLL file not found" );
                case ErrorCodes.FILE_COULDNT_BE_OPEN:
                    throw new Exception( "Couldn't open " + dllPath );
                default:
                    throw new Exception( "Internal Error: DLLType() returned some stupid shit\r\nDLL file: " + dllPath );
            }

            DLLLoader.LoadDLL( dllPath );

            AcceptConnectionRoutine = DLLLoader.GetAcceptConnectionDelegate();
            IsReceiverDelegate = DLLLoader.GetIsReceiverDelegate();
            GetMIDDelegate = DLLLoader.GetGetMIDDelegate();
            MatchMIDDelegate = DLLLoader.GetMatchMIDDelegate();
            GetDataDelegate = DLLLoader.GetGetDataDelegate();
            DispatchDataDelegate = DLLLoader.GetDispatchDataDelegate();
            StopAcceptingRoutine = DLLLoader.GetStopAcceptingDelegate();
            TerminateRoutine = DLLLoader.GetTerminateDelegate();
            CloseConnectionDelegate = DLLLoader.GetCloseConnectionDelegate();

            _IsTerminating = false;
        }
        
        /*
         * Supposed to call a DLL function, that
         * accepts connection and returns an id
         * that the dll uses to identify the entity
         * and then 
         */
        public ExternalEntity AcceptConnection()
        {
            // if we are terminating, quit right now
            if (_IsTerminating)
            {
                return null;
            }

            UInt32 newID = AcceptConnectionRoutine();

            // if we are terminating, quit right now
            if (_IsTerminating)
            {
                return null;
            }

            return new ExternalEntity( newID, IsReceiverDelegate );
        }
        
        public void PassMIDDelegates( IDefinedEntity entity )
        {
            entity.AcceptMIDDelegates( GetMIDDelegate, MatchMIDDelegate );
        }

        public void PassGetDataDelegateToSender( Sender sender )
        {
            sender.AcceptGetDataDelegate( GetDataDelegate );
        }

        public void PassCloseConnectionDelegate( IDefinedEntity entity )
        {
            entity.AcceptCloseConnectionDelegate( CloseConnectionDelegate );
        }

        public void PassDispatchDataDelegateToReceiver( Receiver receiver )
        {
            receiver.AcceptDispatchDataDelegate( DispatchDataDelegate );
        }

        public void CloseAccepting()
        {
            _IsTerminating = true;

            StopAcceptingRoutine();
        }

        public void Dispose()
        {
            TerminateRoutine();

            DLLLoader.UnloadDLL();

            DLLLoader = null;
        }
    }
}
