using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Dicom.Printing
{
    public class Printer : DicomDataset
    {
        public string PrinterAet { get; private set; }

        public string PrinterStatus
        {
            get { return Get<string>(DicomTag.PrinterStatus, "NORMAL"); }
            private set { Add(DicomTag.PrinterStatus, value); }
        }

        public string PrinterStatusInfo
        {
            get { return Get<string>(DicomTag.PrinterStatusInfo, "NORMAL"); }
            private set { Add(DicomTag.PrinterStatusInfo, value); }
        }

        // AHORA: setter público
        public string PrinterName
        {
            get { return Get(DicomTag.PrinterName, string.Empty); }
            set { Add(DicomTag.PrinterName, value); }
        }

        public string Manufacturer
        {
            get { return Get(DicomTag.Manufacturer, "Nebras Technology"); }
            private set { Add(DicomTag.Manufacturer, value); }
        }

        public string ManufacturerModelName
        {
            get { return Get(DicomTag.ManufacturerModelName, "PaXtreme Printer"); }
            private set { Add(DicomTag.ManufacturerModelName, value); }
        }

        public string DeviceSerialNumber
        {
            get { return Get(DicomTag.DeviceSerialNumber, string.Empty); }
            private set { Add(DicomTag.DeviceSerialNumber, value); }
        }

        public string SoftwareVersions
        {
            get { return Get(DicomTag.SoftwareVersions, string.Empty); }
            private set { Add(DicomTag.SoftwareVersions, value); }
        }

        public DateTime DateTimeOfLastCalibration
        {
            get { return this.GetDateTime(DicomTag.DateOfLastCalibration, DicomTag.TimeOfLastCalibration); }
            private set
            {
                Add(DicomTag.DateOfLastCalibration, value);
                Add(DicomTag.TimeOfLastCalibration, value);
            }
        }

        public Printer(string aet)
        {
            PrinterAet = aet;
            DateTimeOfLastCalibration = DateTime.Now;

            PrinterStatus = "NORMAL";
            PrinterStatusInfo = "NORMAL";
        }
    }
}
