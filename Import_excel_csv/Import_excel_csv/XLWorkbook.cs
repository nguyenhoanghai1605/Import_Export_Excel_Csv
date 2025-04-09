using System;
using System.IO;

namespace Tools_LayHinh
{
    internal class XLWorkbook
    {
        private FileStream stream;

        public XLWorkbook(FileStream stream)
        {
            this.stream = stream;
        }

        internal object Worksheet(int v)
        {
            throw new NotImplementedException();
        }
    }
}