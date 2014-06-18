using System;
using Microsoft.PointOfService;
using WebPrint.Framework;

namespace WebPrint.Pos
{
    public class PosPrinter : IDisposable
    {
        private PosExplorer _explorer;
        private Microsoft.PointOfService.PosPrinter _printer;

        public string LogicalName { get; private set; }

        private PosExplorer Explorer
        {
            get { return _explorer ?? (_explorer = new PosExplorer()); }
        }

        private Microsoft.PointOfService.PosPrinter Printer
        {
            get
            {
                if (_printer != null)
                {
                    InitPrinter();
                    return _printer;
                }

                if (LogicalName.IsNullOrEmpty())
                    throw new ArgumentNullException("logical name is null or empty.");

                var device = Explorer.GetDevice(DeviceType.PosPrinter, LogicalName);
                if (device == null)
                    throw new NullReferenceException(
                        "Can't find the device by logicalName : {0}.".Formatting(LogicalName));

                _printer = Explorer.CreateInstance(device) as Microsoft.PointOfService.PosPrinter;

                if (_printer == null)
                    throw new NullReferenceException(
                        "Create the instance of Microsoft.PointOfService.PosPrinter faild.");

                InitPrinter();
                return _printer;
            }
        }

        private void InitPrinter()
        {
            if (_printer.State == ControlState.Closed)
                _printer.Open();

            if (!_printer.Claimed)
                _printer.Claim(0);

            if (!_printer.DeviceEnabled)
                _printer.DeviceEnabled = true;

            if (!_printer.RecLetterQuality)
                // If true, prints in high-quality mode. If false, prints in high-speed mode
                _printer.RecLetterQuality = true;
        }

        public PosPrinter(string logicalName)
        {
            LogicalName = logicalName;
        }

        public void Print(string receipt)
        {
            Printer.AsyncMode = false;

            InternalPrint(receipt);
        }

        public void PrintAsync(string receipt)
        {
            Printer.AsyncMode = true;

            InternalPrint(receipt);
        }

        private void InternalPrint(string receipt)
        {
            receipt = receipt.Replace("ESC", ((char) 27).ToString());

            Printer.PrintNormal(PrinterStation.Receipt, receipt);
        }

        public void Close()
        {
            if (_printer != null) 
                _printer.Close();
        }

        public void Dispose()
        {
            this.Close();
        }
    }
}
