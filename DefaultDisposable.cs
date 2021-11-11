using System;

namespace Rehau.Sku.Assist
{
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

}


