using ExcelDna.Integration;
using System.Threading.Tasks;
using System;
using System.Threading;

namespace Rehau.Sku.Assist
{
    internal static class ExcelTaskUtil
    {
        public static object Run<TResult>(string callerFunctionName, object callerParameters, Func<CancellationToken, Task<TResult>> taskSource)
        {
            return ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, delegate
            {
                var cts = new CancellationTokenSource();
                var task = taskSource(cts.Token);
                return new ExcelTaskObservable<TResult>(task, cts);
            });
        }

        public static object Run<TResult>(string callerFunctionName, object callerParameters, Func<Task<TResult>> taskSource)
        {
            return ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, delegate
            {
                var task = taskSource();
                return new ExcelTaskObservable<TResult>(task);
            });
        }

        class ExcelTaskObservable<TResult> : IExcelObservable
        {
            readonly Task<TResult> _task;
            readonly CancellationTokenSource _cts;
            public ExcelTaskObservable(Task<TResult> task)
            {
                _task = task;
            }

            public ExcelTaskObservable(Task<TResult> task, CancellationTokenSource cts)
                : this(task)
            {
                _cts = cts;
            }

            public IDisposable Subscribe(IExcelObserver observer)
            {
                switch (_task.Status)
                {
                    case TaskStatus.RanToCompletion:
                        observer.OnNext(_task.Result);
                        observer.OnCompleted();
                        break;
                    case TaskStatus.Faulted:
                        observer.OnError(_task.Exception.InnerException);
                        break;
                    case TaskStatus.Canceled:
                        observer.OnError(new TaskCanceledException(_task));
                        break;
                    default:
                        _task.ContinueWith(t =>
                        {
                            switch (t.Status)
                            {
                                case TaskStatus.RanToCompletion:
                                    observer.OnNext(t.Result);
                                    observer.OnCompleted();
                                    break;
                                case TaskStatus.Faulted:
                                    observer.OnError(t.Exception.InnerException);
                                    break;
                                case TaskStatus.Canceled:
                                    observer.OnError(new TaskCanceledException(t));
                                    break;
                            }
                        });
                        break;
                }

                if (_cts != null)
                {
                    return new CancellationDisposable(_cts);
                }

                return DefaultDisposable.Instance;
            }
        }

        sealed class DefaultDisposable : IDisposable
        {

            public static readonly DefaultDisposable Instance = new DefaultDisposable();

            DefaultDisposable()
            {
            }

            public void Dispose()
            {
            }
        }

        sealed class CancellationDisposable : IDisposable
        {

            readonly CancellationTokenSource cts;
            public CancellationDisposable(CancellationTokenSource cts)
            {
                if (cts == null)
                {
                    throw new ArgumentNullException("cts");
                }

                this.cts = cts;
            }

            public CancellationDisposable()
                : this(new CancellationTokenSource())
            {
            }

            public CancellationToken Token
            {
                get { return cts.Token; }
            }

            public void Dispose()
            {
                cts.Cancel();
            }
        }

    }
}

