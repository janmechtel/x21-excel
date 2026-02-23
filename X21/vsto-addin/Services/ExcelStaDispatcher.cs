using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;
using X21.Logging;

namespace X21.Services
{
    /// <summary>
    /// Routes every Excel COM call through a single STA thread to satisfy
    /// Excel's apartment model requirements.
    /// </summary>
    public sealed class ExcelStaDispatcher : IDisposable
    {
        private readonly BlockingCollection<Action> _workQueue = new BlockingCollection<Action>();
        private readonly Thread _staThread;
        private int _staThreadId;
        private bool _disposed;

        public ExcelStaDispatcher()
        {
            _staThread = new Thread(DispatchLoop)
            {
                IsBackground = true,
                Name = "X21.ExcelStaDispatcher"
            };

            _staThread.SetApartmentState(ApartmentState.STA);
            _staThread.Start();
        }

        public void InvokeExcel(Action action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            if (Thread.CurrentThread.ManagedThreadId == _staThreadId)
            {
                action();
                return;
            }

            var tcs = new TaskCompletionSource<object>(TaskCreationOptions.RunContinuationsAsynchronously);
            Enqueue(() =>
            {
                try
                {
                    action();
                    tcs.SetResult(null);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });

            tcs.Task.GetAwaiter().GetResult();
        }

        public T InvokeExcel<T>(Func<T> func)
        {
            Logger.Info($"InvokeExcel<T>: Called on thread {Thread.CurrentThread.ManagedThreadId} (STA thread: {_staThreadId})");
            if (func == null)
            {
                throw new ArgumentNullException(nameof(func));
            }

            if (Thread.CurrentThread.ManagedThreadId == _staThreadId)
            {
                Logger.Info("InvokeExcel<T>: Already on STA thread, executing directly");
                return func();
            }

            Logger.Info("InvokeExcel<T>: Marshalling to STA thread via queue");
            var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            Enqueue(() =>
            {
                try
                {
                    Logger.Info($"InvokeExcel<T>: Executing on STA thread {Thread.CurrentThread.ManagedThreadId}");
                    var result = func();
                    Logger.Info("InvokeExcel<T>: Function execution completed, setting result");
                    tcs.SetResult(result);
                }
                catch (Exception ex)
                {
                    Logger.Info($"InvokeExcel<T>: ❌ Exception during execution: {ex.Message}");
                    Logger.LogException(ex);
                    tcs.SetException(ex);
                }
            });

            Logger.Info("InvokeExcel<T>: Waiting for STA thread to complete");
            var result = tcs.Task.GetAwaiter().GetResult();
            Logger.Info("InvokeExcel<T>: STA thread completed, returning result");
            return result;
        }

        private void Enqueue(Action action)
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(ExcelStaDispatcher));
            }

            _workQueue.Add(action);
        }

        private void DispatchLoop()
        {
            _staThreadId = Thread.CurrentThread.ManagedThreadId;

            foreach (var action in _workQueue.GetConsumingEnumerable())
            {
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    Logger.LogException(ex);
                }
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _workQueue.CompleteAdding();

            if (!_staThread.Join(TimeSpan.FromSeconds(5)))
            {
                _staThread.Interrupt();
            }
        }
    }
}
