using System;
using System.Threading;

namespace Rehau.Sku.Assist
{
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


