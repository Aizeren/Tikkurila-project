using System;
using System.Threading;

namespace SimplifiedCore
{
    /*
     * Every Sender-Receiver connection is
     * sustained by this class - using constructor,
     * we assign it a sender and a receiver.
     * 
     * To open connection, funciton OpenTransfer()
     * is used - it starts the main connection thread,
     * that receives data from sender and then sends
     * it to receiver in a loop.
     * 
     * Use CloseTransfer() function to close the
     * connection.
     * 
     * Function MatchCheck() is called if we want
     * to find out if this is the Connection we are
     * seeking, as an example, when we need to close
     * a connection between certain "points", that
     * we have references for.
     */
    class Connection : IDisposable
    {
        private Sender _Sender;
        private Receiver _Receiver;

        private Thread _TransferThread;

        private CancellationTokenSource _TransferLoop;

        /*
         * The routine, run by the main thread (_TransferThread)
         */
        private void TransferMain( CancellationToken transferLoopToken )
        {
            while (!transferLoopToken.IsCancellationRequested)
            {
                // Rexarrior claimed there was a bug, when this array
                // was being created outside the loop and putting
                // the code line inside fixes it
                byte[] data = new byte[Numerics.BufferSizes.DataTransferBuffer];

                // UNSAFE code here - this freezes, if
                // DLL's GetData or DispatchData hangs
                _Sender.GetData(data);
                // SNIFF_POINT Point, when we can sniff and analyse messages
                _Receiver.DispatchData(data);

                // OPTIONAL - a very short sleep to prevent high CPU loading
                Thread.Sleep(Numerics.Timeouts.DataTransferIterationTimeout);
            }
        }

        public int OpenTransfer()
        {
            // _TransferThread is null only right after
            // creation of the connection, before we run it
            // for the first time
            if (_TransferThread != null) // if this is fresh, skip this
            {
                if (!_TransferLoop.IsCancellationRequested) // if this is not fresh, and cancellation hasn't been requested - it's now running
                {
                    return ErrorCodes.ALREADY_RUNNING;
                }
                /*
                // we can actually put away else and just
                // put its body here, because the "if" above
                // just returns. but "else" makes it clearer,
                // that we run the code below when cancellation
                // has already been requested.
                */
                else
                {
                    /*
                     * okay, cancellation has been requested.
                     * it's good if everything has terminated
                     * by the moment we run OpenTransfer() again.
                     * But if not, the only thing that is probably
                     * hang is the _TransferThread
                     */
                    if (_TransferThread.IsAlive)
                    {
                        if (!_TransferThread.Join(Numerics.Timeouts.OpenTransferWaitForTransferTermination)) // wait 5 seconds
                        {
                            return ErrorCodes.PREVIOUS_RUN_DOESNT_TERMINATE;
                        }
                    }
                }
            }

            // Same as _TransferThread, _TransferLoop
            // is only null, if this is a "fresh" Connection,
            // meaning we haven't started it yet
            if (_TransferLoop != null)
            {
                _TransferLoop.Dispose();
            }
            
            _TransferThread = new Thread( delegate() { TransferMain(_TransferLoop.Token); } );
            _TransferLoop = new CancellationTokenSource();

            _TransferThread.Start();

            return ErrorCodes.ERROR_SUCCESS;
        }

        // TODO: Add safe closing
        // I mean, it is unsafe, because
        // DispatchData and GetData may hang
        // Or not.
        public int CloseTransfer()
        {
            // _TransferThread is null only right after
            // creation of the connection, before we run it
            // for the first time
            if (_TransferThread == null)
            {
                return ErrorCodes.NOT_RUNNING; // haven't started yet
            }
            else
            {
                if (!_TransferThread.IsAlive)
                {
                    return ErrorCodes.NOT_RUNNING; // already terminated
                }
                else
                {
                    if (_TransferLoop.IsCancellationRequested)
                    {
                        return ErrorCodes.PENDING_TERMINATION; // already asked for termination
                    }
                }
            }

            _TransferLoop.Cancel();

            return ErrorCodes.ERROR_SUCCESS;
        }

        public bool MatchCheck(Sender sender, Receiver receiver)
        {
            return ((sender == _Sender) && (receiver == _Receiver));
        }

        /*
         * This function is needed when you
         * have to be sure that the transfer
         * has stopped before moving on
         */
        public int WaitForTermination()
        {
            if (_TransferThread == null)
            {
                return ErrorCodes.NOT_RUNNING;
            }
            else
            {
                if (!_TransferLoop.IsCancellationRequested)
                {
                    return ErrorCodes.TERMINATION_NOT_REQUESTED;
                }
                else
                {
                    _TransferThread.Join();
                    return ErrorCodes.ERROR_SUCCESS;
                }
            }
        }

        public void CloseBothConnections()
        {
            // Close connections point -
            // here we call .dll functions
            _Sender.CloseConnection();
            _Receiver.CloseConnection();
        }

        public void Dispose()
        {
            _TransferThread = null;

            _TransferLoop.Dispose();

            _Sender = null;
            _Receiver = null;
        }

        public Connection(Sender sender, Receiver receiver)
        {
            _Sender = sender;
            _Receiver = receiver;

            _TransferThread = null;

            _TransferLoop = null;
        }
    }
}
