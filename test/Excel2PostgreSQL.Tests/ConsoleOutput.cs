using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2PostgreSQL.Tests
{
    internal sealed class ConsoleOutput : IDisposable
    {
        private volatile bool _disposed;
        private readonly StringWriter _outWriter;
        private readonly StringWriter _errWriter;
        private readonly TextWriter _outBackup;
        private readonly TextWriter _errBackup;

        private ConsoleOutput(bool captureStdOut = true, bool captureStdErr = true)
        {
            if (captureStdOut)
            {
                _outWriter = new StringWriter();
                _outBackup = Console.Out;
                Console.SetOut(_outWriter);
            }
            if (captureStdErr)
            {
                _errWriter = new StringWriter();
                _errBackup = Console.Error;
                Console.SetError(_errWriter);
            }
        }

        public static ConsoleOutput StartCapturing(bool captureStdOut = true, bool captureStdErr = true)
        {
            return new ConsoleOutput();
        }

        public string GetStdErrText()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(ConsoleOutput));
            }
            Console.Error.Flush();
            return _errWriter?.ToString();
        }

        public string GetStdOutText()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(ConsoleOutput));
            }
            Console.Out.Flush();
            return _outWriter?.ToString();
        }

        void IDisposable.Dispose()
        {
            if (_disposed)
            {
                return;
            }
            _disposed = true;
            if (_outWriter != null)
            {
                Console.SetOut(_outBackup);
                _outWriter.Dispose();
            }
            if (_errWriter != null)
            {
                Console.SetError(_errBackup);
                _errWriter.Dispose();
            }
        }
    }
}
