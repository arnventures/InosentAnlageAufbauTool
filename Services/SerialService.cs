using System;
using System.IO.Ports;
using System.Threading;
using System.Threading.Tasks;
using NModbus;
using NModbus.Serial;
using InosentAnlageAufbauTool.Helpers; // ILogger

namespace InosentAnlageAufbauTool.Services
{
    public class SerialService : IDisposable
    {
        private readonly IModbusFactory _factory = new ModbusFactory();
        private readonly ILogger? _logger;

        private IModbusSerialMaster? _master;
        private SerialPort? _serialPort;

        private CancellationTokenSource? _wdCts;
        private Task? _watchdogTask;

        private readonly SemaphoreSlim _gate = new(1, 1); // serialize all bus calls
        private readonly object _stateLock = new();

        private bool _isConnected;
        private bool _disposed;

        // Defaults – tuned for quicker loops
        private const int DefaultReadTimeout = 1000;
        private const int DefaultWriteTimeout = 1000;
        private const int DefaultRetries = 1;

        // Common registers
        public const ushort REG_DEVICE_TYPE = 1;
        public const ushort REG_PRESENCE_CHECK = 2; // any readable register works; 2 = used for quick checks
        public const ushort REG_SERIAL = 3;
        public const ushort REG_SET_ADDRESS = 4;
        public const ushort REG_REBOOT = 17;

        // LED block (write multiple @ 4..7)
        public const ushort REG_LED_MODE = 4;       // optional (0 keeps default)
        public const ushort REG_LED_ADDR = 5;
        public const ushort REG_LED_BAUD = 6;       // e.g. 9600
        public const ushort REG_LED_KEY = 7;       // must be 0x8F8F (36751)

        public bool IsConnected => _isConnected;
        public string? LastError { get; private set; }

        public SerialService(ILogger? logger = null) => _logger = logger;

        // ---------- Connect / Disconnect ----------

        public bool Connect(string portName)
        {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(portName))
            {
                LastError = "Port name is empty.";
                Log("Connect aborted: empty port name.");
                return false;
            }

            lock (_stateLock)
            {
                Log($"Connect requested: {portName}");

                InternalStopWatchdog_NoThrow();
                InternalClose_NoThrow();

                try
                {
                    _serialPort = new SerialPort(portName, 9600, Parity.None, 8, StopBits.One)
                    {
                        ReadTimeout = DefaultReadTimeout,
                        WriteTimeout = DefaultWriteTimeout,
                        DtrEnable = true,
                        RtsEnable = true,
                        Handshake = Handshake.None
                    };

                    Log($"Opening {portName} ...");
                    _serialPort.Open();
                    Log($"{portName} opened.");

                    _master = _factory.CreateRtuMaster(_serialPort);
                    _master.Transport.Retries = DefaultRetries;
                    _master.Transport.ReadTimeout = DefaultReadTimeout;
                    _master.Transport.WriteTimeout = DefaultWriteTimeout;

                    Log("Modbus RTU master created.");

                    _isConnected = true;

                    _wdCts = new CancellationTokenSource();
                    _watchdogTask = Task.Run(() => WatchdogAsync(_wdCts.Token));
                    Log("Watchdog started.");

                    LastError = null;
                    return true;
                }
                catch (Exception ex)
                {
                    LastError = ex.Message;
                    _isConnected = false;
                    Log($"Connect FAILED on {portName}: {ex.GetType().Name}: {ex.Message}");
                    InternalClose_NoThrow();
                    return false;
                }
            }
        }

        public void Disconnect()
        {
            ThrowIfDisposed();
            lock (_stateLock)
            {
                Log("Disconnect()");
                InternalStopWatchdog_NoThrow();
                InternalClose_NoThrow();
                _isConnected = false;
            }
        }

        // ---------- Regular async wrappers used by VM ----------

        public async Task<ushort[]> ReadHoldingAsync(ushort address, ushort count = 1, byte unit = 1)
        {
            ThrowIfDisposed();
            EnsureOpen();

            await _gate.WaitAsync().ConfigureAwait(false);
            try
            {
                Log($"ReadHolding u:{unit} addr:{address} cnt:{count}");
                var data = _master!.ReadHoldingRegisters(unit, address, count); // sync under gate
                return data;
            }
            catch (Exception ex)
            {
                LastError = ex.Message;
                Log($"ReadHolding FAILED: {ex.GetType().Name}: {ex.Message}");
                throw;
            }
            finally
            {
                _gate.Release();
            }
        }

        public async Task WriteSingleAsync(ushort address, ushort value, byte unit = 1)
        {
            ThrowIfDisposed();
            EnsureOpen();

            await _gate.WaitAsync().ConfigureAwait(false);
            try
            {
                Log($"WriteSingle u:{unit} addr:{address} val:{value}");
                _master!.WriteSingleRegister(unit, address, value);
            }
            catch (Exception ex)
            {
                LastError = ex.Message;
                Log($"WriteSingle FAILED: {ex.GetType().Name}: {ex.Message}");
                throw;
            }
            finally
            {
                _gate.Release();
            }
        }

        public async Task WriteMultipleAsync(byte unit, ushort startRegister, ushort[] values)
        {
            ThrowIfDisposed();
            EnsureOpen();

            await _gate.WaitAsync().ConfigureAwait(false);
            try
            {
                Log($"WriteMultiple u:{unit} start:{startRegister} len:{values?.Length ?? 0}");
                _master!.WriteMultipleRegisters(unit, startRegister, values);
            }
            catch (Exception ex)
            {
                LastError = ex.Message;
                Log($"WriteMultiple FAILED: {ex.GetType().Name}: {ex.Message}");
                throw;
            }
            finally
            {
                _gate.Release();
            }
        }

        // ---------- FAST helpers (tight polling, 0 retries) ----------

        private T WithTransport<T>(Func<T> action, int readTimeout, int writeTimeout, int retries)
        {
            EnsureOpen();
            var trans = _master!.Transport;
            int oldRead = trans.ReadTimeout;
            int oldWrite = trans.WriteTimeout;
            int oldRetries = trans.Retries;
            try
            {
                trans.ReadTimeout = readTimeout;
                trans.WriteTimeout = writeTimeout;
                trans.Retries = retries;
                return action();
            }
            finally
            {
                try { trans.ReadTimeout = oldRead; } catch { }
                try { trans.WriteTimeout = oldWrite; } catch { }
                try { trans.Retries = oldRetries; } catch { }
            }
        }

        public bool CheckFast(byte unit, int timeoutMs, CancellationToken ct)
        {
            ThrowIfDisposed();
            EnsureOpen();
            ct.ThrowIfCancellationRequested();

            _gate.Wait(ct);
            try
            {
                return WithTransport(() =>
                {
                    try { return _master!.ReadHoldingRegisters(unit, REG_PRESENCE_CHECK, 1).Length == 1; }
                    catch { return false; }
                }, timeoutMs, timeoutMs, retries: 0);
            }
            catch { return false; }
            finally { _gate.Release(); }
        }

        public ushort ReadRegisterFast(byte unit, ushort address, int timeoutMs, CancellationToken ct)
        {
            ThrowIfDisposed();
            EnsureOpen();
            ct.ThrowIfCancellationRequested();

            _gate.Wait(ct);
            try
            {
                return WithTransport(() =>
                {
                    var regs = _master!.ReadHoldingRegisters(unit, address, 1);
                    return regs.Length == 1 ? regs[0] : (ushort)0;
                }, timeoutMs, timeoutMs, retries: 0);
            }
            finally
            {
                _gate.Release();
            }
        }

        public int ReadSerialFast(byte unit, int timeoutMs, CancellationToken ct)
            => ReadRegisterFast(unit, REG_SERIAL, timeoutMs, ct);

        // Convenience single-register ops used by higher-level logic (fast path)
        public void SetAddress(byte currentUnit, byte nextUnit, CancellationToken ct)
        {
            ThrowIfDisposed();
            EnsureOpen();
            ct.ThrowIfCancellationRequested();

            _gate.Wait(ct);
            try
            {
                Log($"SetAddress {currentUnit} -> {nextUnit}");
                _master!.WriteSingleRegister(currentUnit, REG_SET_ADDRESS, nextUnit);
            }
            finally
            {
                _gate.Release();
            }
        }

        public void SoftRestart(byte unit, CancellationToken ct)
        {
            ThrowIfDisposed();
            EnsureOpen();
            ct.ThrowIfCancellationRequested();

            _gate.Wait(ct);
            try
            {
                Log($"SoftRestart u:{unit}");
                _master!.WriteSingleRegister(unit, REG_REBOOT, 42330);
            }
            finally
            {
                _gate.Release();
            }
        }

        // ---------- LED helper (write 4..7 in one 0x10 frame, includes security key) ----------

        /// <summary>
        /// Writes LED configuration in one block (regs 4..7): mode, newAddr, baud, key(0x8F8F).
        /// Typical: mode=0, baud=9600, key=0x8F8F (36751).
        /// </summary>
        public void LedWriteAddressBaudWithKey(byte unit, byte newAddr, ushort baud = 9600, ushort mode = 0, ushort key = 0x8F8F)
        {
            ThrowIfDisposed();
            EnsureOpen();

            _gate.Wait();
            try
            {
                Log($"LED Write block u:{unit} -> addr:{newAddr}, baud:{baud}, key:{key}");
                ushort[] values = { mode, newAddr, baud, key };
                _master!.WriteMultipleRegisters(unit, REG_LED_MODE, values);
            }
            finally
            {
                _gate.Release();
            }
        }

        // ---------- Utilities ----------

        public void WriteMultiple(byte unit, ushort startRegister, ushort[] values)
        {
            ThrowIfDisposed();
            EnsureOpen();

            _gate.Wait();
            try
            {
                Log($"WriteMultiple (sync) u:{unit} start:{startRegister} len:{values?.Length ?? 0}");
                _master!.WriteMultipleRegisters(unit, startRegister, values);
            }
            finally
            {
                _gate.Release();
            }
        }

        public void FlushBuffers()
        {
            try { _serialPort?.DiscardInBuffer(); } catch { }
            try { _serialPort?.DiscardOutBuffer(); } catch { }
        }

        // ---------- Internals ----------

        private void EnsureOpen()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(SerialService));
            if (!_isConnected) throw new InvalidOperationException("Not connected.");

            if (_serialPort is { IsOpen: false })
            {
                try
                {
                    Log("EnsureOpen: port closed, trying to reopen...");
                    _serialPort.Open();
                    Log("EnsureOpen: reopen success.");
                }
                catch (Exception ex)
                {
                    LastError = ex.Message;
                    Log($"EnsureOpen: reopen FAILED: {ex.GetType().Name}: {ex.Message}");
                }
            }

            if (_serialPort is null || !_serialPort.IsOpen)
            {
                LastError = "Port not open.";
                throw new InvalidOperationException("Port not open.");
            }
        }

        private async Task WatchdogAsync(CancellationToken token)
        {
            try
            {
                while (!token.IsCancellationRequested)
                {
                    await Task.Delay(5000, token).ConfigureAwait(false);

                    lock (_stateLock)
                    {
                        if (_disposed || !_isConnected) return;

                        if (_serialPort is { IsOpen: false })
                        {
                            try
                            {
                                Log("Watchdog: port closed, trying to reopen...");
                                _serialPort.Open();
                                Log("Watchdog: reopen success.");
                            }
                            catch (Exception ex)
                            {
                                LastError = ex.Message;
                                Log($"Watchdog: reopen FAILED: {ex.GetType().Name}: {ex.Message}");
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                Log("Watchdog: canceled.");
            }
            catch (Exception ex)
            {
                Log($"Watchdog: unexpected error: {ex.GetType().Name}: {ex.Message}");
            }
        }

        private void InternalStopWatchdog_NoThrow()
        {
            try
            {
                if (_wdCts != null) Log("Stopping watchdog...");
                _wdCts?.Cancel();
                _watchdogTask?.Wait(300);
            }
            catch { /* ignore */ }
            finally
            {
                _wdCts?.Dispose();
                _wdCts = null;
                _watchdogTask = null;
            }
        }

        private void InternalClose_NoThrow()
        {
            try
            {
                try { _master?.Dispose(); Log("Master disposed."); } catch { }
                _master = null;

                if (_serialPort is not null)
                {
                    try
                    {
                        if (_serialPort.IsOpen)
                        {
                            Log($"Closing {_serialPort.PortName} ...");
                            _serialPort.Close();
                            Log("Port closed.");
                        }
                    }
                    catch { /* ignore */ }
                    try { _serialPort.Dispose(); Log("Port disposed."); } catch { }
                    _serialPort = null;
                }
            }
            catch { /* ignore */ }
        }

        private void ThrowIfDisposed()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(SerialService));
        }

        public void Dispose()
        {
            if (_disposed) return;
            lock (_stateLock)
            {
                if (_disposed) return;
                Log("Dispose()");
                InternalStopWatchdog_NoThrow();
                InternalClose_NoThrow();
                _gate.Dispose();
                _disposed = true;
                GC.SuppressFinalize(this);
            }
        }

        private void Log(string message) => _logger?.Log($"[Serial] {message}");
    }
}
