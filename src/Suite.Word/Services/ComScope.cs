using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Suite.Word.Services
{
    internal sealed class ComScope : IDisposable
    {
        private readonly List<object> _objects = new List<object>();

        public T Track<T>(T value)
        {
            if (value != null && Marshal.IsComObject(value)) _objects.Add(value);
            return value;
        }

        public void Dispose()
        {
            for (int index = _objects.Count - 1; index >= 0; index--)
            {
                try { Marshal.FinalReleaseComObject(_objects[index]); }
                catch (InvalidComObjectException) { }
            }
            _objects.Clear();
        }
    }
}
