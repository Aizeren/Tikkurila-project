using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.IO;

/*
 * commentary notes to check out:
 * 
 * NEEDTOKNOW - comments with this note ABSOLUTELY
 * HAVE TO BE READ AND UNDERSTOOD by those, who are
 * gonna work somehow with the system (Core).
 * 
 * UNSAFE - (it's not about the code being unmanaged,
 * it's exatcly about it being unsafe) - make it safer
 * TODO - implement/rework
 * SNIFF_POINT - where we can mitm the messages
 * OPTIONAL - this may be omitted (especially consider it in Release versions)
 * 
 * BALANCE - make code (or comments) look more "balanced" and convenient
 */

    /*
     * TODO:
     * Make all the function explanations
     * use <summary>, so that we could
     * see the description just hovering
     * our mouse cursor over it.
     */
namespace SimplifiedCore
{
    /*
     * When an external entity connects to the core,
     * it is "registered" - i.e put in a corresponding list -
     * either as a receiver (put in list of receivers) or
     * as a sender (put in list of senders).
     * But before getting it into a list, we check if
     * there is a matching entity in the opposite list.
     * If we found one, the current entity is even not gonna
     * get put in its list, because there is no need of it - 
     * no one will connect to it when it is already in a connection.
     * 
     * To match a corresponding entity in a list, MatchMID() function
     * is used. The MID is retreived from a connecting entity using
     * GetMID() function. MID stands for "Match ID" - you see, the ID,
     * that is used to Match.     
     * A VERY IMPORTANT NOTICE is that if we found a one-sided match,
     * like receiver.MatchMID( sender.GetMID() ) returned true,
     * WE DO NOT ESTABLISH THE CONNECTION BETWEEN THEM UNTIL
     * OPPOSITE MATCH IS PERFORMED - speaking about this example,
     * it would be sender.MatchMID( receiver.GetMID ) - this must return true
     * or otherwise we skip the pair (and maybe log this event
     * in a base of interesting events)
     * 
     * CONFIG FILE:
     * Legend:
     * [optinal] - this can be omitted
     * this/that/those/.. - only one must present
     * <template> - this must be replaced with the
     * string that corresponds the word in those
     * angle brackets
     * ... - "and so on" - this 
     *
     * Notes:
     * All the paths, given in the file, that are not full,
     * are evaluated starting with the curent directory.
     * But the very current directory may be specified
     * differently. This is given in the first line of the
     * config file.
     * 
     * File Format:
     * first line - where the current directory is located: configfile/executable/<directory>
     * if it is "configfile", then current directory is set to the location
     * of the config file, same with "executable" - where the executable is located.
     * If none on those met, we try to set it to the given string,
     * assuming it specifies some directory. If we don't succeed,
     * just return an error PATH_NOT_FOUND
     * Current directory is needed to calculate each of the next
     * paths, if they are not full.
     * Next strings just locate the .dll files used as interfaces:
     * [<.dll path>]
     * [<.dll path>]
     * ...
     * 
     * CODE REVIEW POSSIBLE:
     * When we pass IDs or delegates to the
     * connected entities, we use that strange
     * mechanism: Pass{ entity.Accept( field ) },
     * just because we don't want Core to know
     * anything about those fields. But it looks
     * pretty awkward.
     */
    class Core
    {
        // in this tuple we save externalinterface object
        // and the thread, that calls AcceptConnection for this
        // interface in loop
        private List<Tuple<ExternalInterface, Thread>> _ExternalInterfaces;

        private List<Connection> _Connections;

        private List<Receiver> _WaitingReceivers;
        private List<Sender> _WaitingSenders;

        private CancellationTokenSource _AcceptConnectionsLoop;

        private string _CurrentDirectory;

        // because many connections may be accepted
        // simultaneously, we need to control access
        // to the Lists (above) -> so we do lock(<The SyncRoot Below>)
        private object _NewConnectionSyncRoot = new object();

        /*
            this syncroot is used to hang Start() function
            until everything from the previous time we
            run the Core is disposed. this is possibly
            needed when we call Stop(false) - so we don't
            wait until all the resources are disposed, but
            before the next Start(), we actually NEED them
            to be released
        */
        private object _CleanUpSyncRoot = new object();
        private Thread _CleanUpThread;

        private void AddTransferConnection(Sender sender, Receiver receiver)
        {
            _Connections.Add(new Connection(sender, receiver));
            _Connections.Last().OpenTransfer();
        }

        private void RegisterReceiver(Receiver receiver)
        {
            for (int i = 0; i < _WaitingSenders.Count; i++)
            {
                if (_WaitingSenders[i].MatchMID(receiver.GetMID()))
                {
                    // both-sided check
                    if (!receiver.MatchMID(_WaitingSenders[i].GetMID()))
                    {
                        continue;
                    }

                    AddTransferConnection(_WaitingSenders[i], receiver);
                    _WaitingSenders.RemoveAt(i);
                    return;
                }
            }

            // matching sender not found

            _WaitingReceivers.Add(receiver);
        }

        private void RegisterSender(Sender sender)
        {
            for (int i = 0; i < _WaitingReceivers.Count; i++)
            {
                if (_WaitingReceivers[i].MatchMID(sender.GetMID()))
                {
                    // both-sided check
                    if (!sender.MatchMID(_WaitingReceivers[i].GetMID()))
                    {
                        continue;
                    }

                    AddTransferConnection(sender, _WaitingReceivers[i]);
                    _WaitingReceivers.RemoveAt(i);
                    return;
                }
            }

            // matching receiver not found

            _WaitingSenders.Add(sender);
        }

        /*
         * This is only used to remove
         * Connections while the whole
         * system is runing.
         */
        private void RemoveTransferConnection(Sender sender, Receiver receiver)
        {
            lock (_NewConnectionSyncRoot)
            {
                for (int i = 0; i < _Connections.Count(); i++)
                {
                    if (_Connections[i].MatchCheck(sender, receiver))
                    {
                        _Connections[i].CloseTransfer();
                        _Connections.RemoveAt(i);
                        i--; // assuming that there may be more than one match
                    }
                }
            }
        }

        private void AcceptConnection(ExternalInterface externalInterface)
        {
            ExternalEntity newComer = externalInterface.AcceptConnection();

            if (newComer == null)
            {
                return;
            }

            lock (_NewConnectionSyncRoot)
            {
                if (newComer.IsReceiver())
                {
                    Receiver newReceiver = new Receiver();

                    newComer.PassID(newReceiver);
                    externalInterface.PassMIDDelegates(newReceiver);
                    externalInterface.PassDispatchDataDelegateToReceiver(newReceiver);
                    externalInterface.PassCloseConnectionDelegate(newReceiver);

                    RegisterReceiver(newReceiver);
                }
                else
                {
                    Sender newSender = new Sender();

                    newComer.PassID(newSender);
                    externalInterface.PassMIDDelegates(newSender);
                    externalInterface.PassGetDataDelegateToSender(newSender);
                    externalInterface.PassCloseConnectionDelegate(newSender);

                    RegisterSender(newSender);
                }

            }
        }

        private void AcceptConnectionsMain( ExternalInterface externalInterface, CancellationToken acceptConnectionsLoopToken )
        {
            while (!acceptConnectionsLoopToken.IsCancellationRequested)
            {
                AcceptConnection(externalInterface);
            }
        }

        private void RegisterInterface(String dllPath)
        {
            ExternalInterface externalInterface = null;

            try
            {
                 externalInterface = new ExternalInterface(MakePath(dllPath));
            }
            catch ( Exception ex )
            {
                Console.WriteLine( ex.Message );
                return;
            }

            Thread interfaceThread = new Thread(
                delegate ()
                {
                    AcceptConnectionsMain( externalInterface, _AcceptConnectionsLoop.Token );
                }
                );

            interfaceThread.Start();

            _ExternalInterfaces.Add( new Tuple<ExternalInterface, Thread>(externalInterface, interfaceThread) );
        }

        /*
         * This function is used as a
         * Thread startup routine for
         * _CleanUpThread
         */
        private void CleanUp()
        {
            // first thing we should do is to pervent accepting
            // new entities and adding them into corresponding lists

            // UNSAFE, because ExternalInterface doesn't handle
            // exceptions, that may be raised when we interrupt
            // connections to those waiting receivers and senders
            // OR NOT UNSAFE because we call CloseConnection() routines
            // from all the Receivers and Senders?

            foreach (Tuple<ExternalInterface, Thread> externalInterface in _ExternalInterfaces)
            {
                externalInterface.Item1.CloseAccepting();
                externalInterface.Item2.Join();
            }

            // its tokens are used only in the threads
            // that accept external connections in
            // the loop. When all the threads are terminated,
            // we can Dispose the CancellationTokenSource
            _AcceptConnectionsLoop.Dispose();

            // now we need to close all the connections

            foreach (Connection connection in _Connections)
            {
                connection.CloseTransfer();
                // ya, we CloseConnections before Wait(ing)ForTermination
                // that's because GetData() and DispatchData() in Connection
                // may hang
                connection.CloseBothConnections();
                // wait for Connection main thread
                // to terminate properly
                connection.WaitForTermination();
                // finally, we can Dispose all the
                // resources that were allocated
                connection.Dispose();
            }

            // all the connections are closed.
            // clear the list.
            _Connections.Clear();

            // remove all the remaining entities
            foreach (Receiver receiver in _WaitingReceivers)
            {
                receiver.CloseConnection();
            }

            foreach (Sender sender in _WaitingSenders)
            {
                sender.CloseConnection();
            }

            _WaitingReceivers.Clear();
            _WaitingSenders.Clear();

            foreach (Tuple<ExternalInterface, Thread> externalInterface in _ExternalInterfaces)
            {
                externalInterface.Item1.Dispose();
            }

            // Now, nothing vivid left in this list -
            // so, clear it
            _ExternalInterfaces.Clear();
        }

        /*
         * Every path found in config file
         * must be passed through this function,
         * execpt the one, that is at the first line,
         * if the first line represents a path of course.
         * 
         * So, this is gonna be like "currentString = MakePath( currentString )"
         */
        private string MakePath(string givenPath)
        {
            // Full path are complete
            if (Path.IsPathRooted(givenPath))
            {
                return givenPath;
            }

            return Path.Combine(_CurrentDirectory, givenPath);
        }

        /*
         * TODO: We do not handle the situation,
         * when a single DLL met more than once
         * in a config file BUT MAYBE... we don't
         * even need to handle this case...
         */
        public int Start(String configFilePath)
        {
            // supposing now, that there are only DLL paths
            // in the config file


            if (_AcceptConnectionsLoop != null && !_AcceptConnectionsLoop.IsCancellationRequested)
            {
                return ErrorCodes.ALREADY_RUNNING;
            }

            // wait for the previous termination
            if (_CleanUpThread != null)
            {
                _CleanUpThread.Join();
            }

            // Create CancellationToken Exactly Here
            _AcceptConnectionsLoop = new CancellationTokenSource();

            if (!File.Exists(configFilePath))
            {
                return ErrorCodes.FILE_NOT_FOUND;
            }

            FileStream configFile = new FileStream(configFilePath, FileMode.Open, FileAccess.Read, FileShare.None);
            StreamReader configReader = new StreamReader(configFile);

            String currentString;

            currentString = configReader.ReadLine();
            // first actually we should check if configFilePath is rooted
            if (currentString == Strings.ConfigCurrentDirectory.ConfigFile)
            {
                _CurrentDirectory = Path.GetDirectoryName(configFilePath);
            }
            else
            {
                if (currentString == Strings.ConfigCurrentDirectory.Executable)
                {
                    _CurrentDirectory = Directory.GetCurrentDirectory();
                }
                else
                {
                    if (currentString == Strings.ConfigCurrentDirectory.Path)
                    {
                        currentString = configReader.ReadLine();
                        if (Directory.Exists(currentString))
                        {
                            _CurrentDirectory = currentString;
                        }
                        else
                        {
                            return ErrorCodes.PATH_NOT_FOUND;
                        }
                    }
                }
            }

            while (!configReader.EndOfStream)
            {
                currentString = configReader.ReadLine();
                RegisterInterface(currentString);
            }

            configReader.Close();
            configFile.Close();

            return ErrorCodes.ERROR_SUCCESS;
        }

        /*
         * If waitForCompleteTermination is false, the function
         * quits, leaving _CleanUpThread running. If it's not,
         * it waits for _CleanUpThread to terminate as many
         * milliseconds as specified by the second parameter.
         * If waitMillisecs is zero, the waiting can only be
         * finished by _CleanUpThread terminating.
         * So that, waitMillisecs matters if only the first
         * parameter is true. Since it is optional and its
         * default value is 0, if you want to wait for certain
         * completion, just call Stop(true).
         */
        public int Stop(bool waitForCompleteTermination, int waitMillisecs = 0)
        {
            if (_AcceptConnectionsLoop == null)
            {
                return ErrorCodes.NOT_RUNNING;
            }
            else
            {
                if (_AcceptConnectionsLoop.IsCancellationRequested)
                {
                    if (_CleanUpThread.IsAlive)
                    {
                        return ErrorCodes.PENDING_TERMINATION; // BALANCE not exactly "pending", because _CleanUpThread terminates the Core
                    }
                    else
                    {
                        return ErrorCodes.NOT_RUNNING;
                    }
                }
            }

            _AcceptConnectionsLoop.Cancel();

            _CleanUpThread = new Thread( CleanUp );

            _CleanUpThread.Start();

            if (waitForCompleteTermination)
            {
                if (waitMillisecs == 0)
                {
                    _CleanUpThread.Join();
                }
                else
                {
                    _CleanUpThread.Join( waitMillisecs );
                }
            }

            return ErrorCodes.ERROR_SUCCESS;
        }

        public Core()
        {
            _Connections = new List<Connection>();

            _WaitingReceivers = new List<Receiver>();
            _WaitingSenders = new List<Sender>();

            _ExternalInterfaces = new List<Tuple<ExternalInterface, Thread>>();

            _AcceptConnectionsLoop = null;
        }
    }
}
