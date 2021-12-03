using System;

namespace Rehau.Sku.Assist
{
    interface IProduct
    {
        string Sku { get; }
        string Name { get; }
        Uri Uri { get; }
    }
}
