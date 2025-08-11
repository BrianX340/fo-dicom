using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Linq;
using System.Reflection;
using System.Threading;
using Dicom;
using Dicom.Imaging;

namespace Dicom.Printing
{
    public class StatusUpdateEventArgs : EventArgs
    {
        public ushort EventTypeId { get; private set; }
        public string ExecutionStatusInfo { get; private set; }
        public string FilmSessionLabel { get; private set; }
        public string PrinterName { get; private set; }

        public StatusUpdateEventArgs(ushort eventTypeId, string executionStatusInfo, string filmSessionLabel, string printerName)
        {
            EventTypeId = eventTypeId;
            ExecutionStatusInfo = executionStatusInfo;
            FilmSessionLabel = filmSessionLabel;
            PrinterName = printerName;
        }
    }

    public enum PrintJobStatus : ushort
    {
        Pending = 1,
        Printing = 2,
        Done = 3,
        Failure = 4
    }

    public class PrintJob : DicomDataset
    {
        public bool SendNEventReport { get; set; }
        private readonly object _synchRoot = new object();

        public Guid PrintJobGuid { get; private set; }
        public IList<string> FilmBoxFolderList { get; private set; }
        public Printer Printer { get; private set; }
        public PrintJobStatus Status { get; private set; }
        public string PrintJobFolder { get; private set; }
        public string FullPrintJobFolder { get; private set; }
        public Exception Error { get; private set; }
        public string FilmSessionLabel { get; private set; }

        // Preferencias por ruta
        public Duplex? Duplex { get; set; }
        public string ForcedPaperSize { get; set; }     // "A4" | "Letter" | "10INX12IN" | "24CMX30CM"
        public string ForcedPaperSource { get; set; }   // "Bandeja 2" | "Tray 2" | etc.
        public bool FitToPage { get; set; } = true;

        private int _currentPage;
        private FilmBox _currentFilmBox;
        private bool _polarityReverse;

        // Señal para responder OK cuando el trabajo ya está en cola
        private readonly ManualResetEventSlim _queuedEvt = new ManualResetEventSlim(false);
        public bool WaitUntilQueued(TimeSpan timeout) => _queuedEvt.Wait(timeout);

        public readonly DicomUID SOPClassUID = DicomUID.PrintJobSOPClass;
        public DicomUID SOPInstanceUID { get; private set; }

        public string ExecutionStatus { get { return Get(DicomTag.ExecutionStatus, string.Empty); } set { Add(DicomTag.ExecutionStatus, value); } }
        public string ExecutionStatusInfo { get { return Get(DicomTag.ExecutionStatusInfo, string.Empty); } set { Add(DicomTag.ExecutionStatusInfo, value); } }
        public string PrintPriority { get { return Get(DicomTag.PrintPriority, "MED"); } set { Add(DicomTag.PrintPriority, value); } }

        public DateTime CreationDateTime
        {
            get { return this.GetDateTime(DicomTag.CreationDate, DicomTag.CreationTime); }
            set { Add(DicomTag.CreationDate, value); Add(DicomTag.CreationTime, value); }
        }

        public string PrinterName { get { return Get(DicomTag.PrinterName, string.Empty); } set { Add(DicomTag.PrinterName, value); } }
        public string Originator { get { return Get(DicomTag.Originator, string.Empty); } set { Add(DicomTag.Originator, value); } }

        public Dicom.Log.Logger Log { get; private set; }
        public event EventHandler<StatusUpdateEventArgs> StatusUpdate;

        public PrintJob(DicomUID sopInstance, Printer printer, string originator, Dicom.Log.Logger log)
            : base()
        {
            if (printer == null) throw new ArgumentNullException("printer");
            Log = log;

            SOPInstanceUID = (sopInstance == null || sopInstance.UID == string.Empty) ? DicomUID.Generate() : sopInstance;
            this.Add(DicomTag.SOPClassUID, SOPClassUID);
            this.Add(DicomTag.SOPInstanceUID, SOPInstanceUID);

            Printer = printer;
            Status = PrintJobStatus.Pending;

            PrinterName = Printer.PrinterName;
            Originator = originator;

            if (CreationDateTime == DateTime.MinValue) CreationDateTime = DateTime.Now;

            PrintJobFolder = SOPInstanceUID.UID;
            var receivingFolder = Environment.CurrentDirectory + @"\PrintJobs";
            FullPrintJobFolder = string.Format(@"{0}\{1}", receivingFolder.TrimEnd('\\'), PrintJobFolder);

            FilmBoxFolderList = new List<string>();
        }

        public void Print(IList<FilmBox> filmBoxList)
        {
            try
            {
                Status = PrintJobStatus.Pending;
                OnStatusUpdate("Preparing films for printing");

                var printJobDir = new System.IO.DirectoryInfo(FullPrintJobFolder);
                if (!printJobDir.Exists) printJobDir.Create();

                DicomFile file;
                int filmsCount = FilmBoxFolderList.Count;
                for (int i = 0; i < filmBoxList.Count; i++)
                {
                    var filmBox = filmBoxList[i];
                    var filmBoxDir = printJobDir.CreateSubdirectory(string.Format("F{0:000000}", i + 1 + filmsCount));

                    file = new DicomFile(filmBox.FilmSession);
                    file.Save(string.Format(@"{0}\FilmSession.dcm", filmBoxDir.FullName));

                    FilmBoxFolderList.Add(filmBoxDir.Name);
                    filmBox.Save(filmBoxDir.FullName);
                }

                FilmSessionLabel = filmBoxList.First().FilmSession.FilmSessionLabel;

                var thread = new Thread(new ThreadStart(DoPrint));
                thread.Name = string.Format("PrintJob {0}", SOPInstanceUID.UID);
                thread.IsBackground = true;
                thread.Start();
            }
            catch (Exception ex)
            {
                Error = ex;
                Status = PrintJobStatus.Failure;
                OnStatusUpdate("Print failed");
                DeletePrintFolder();
            }
        }

        private void DoPrint()
        {
            PrintDocument printDocument = null;
            try
            {
                Status = PrintJobStatus.Printing;
                OnStatusUpdate("Printing Started");

                var printerSettings = new PrinterSettings { PrinterName = Printer.PrinterName, PrintToFile = false };
                if (Duplex.HasValue) printerSettings.Duplex = Duplex.Value;

                printDocument = new PrintDocument
                {
                    PrinterSettings = printerSettings,
                    DocumentName = Thread.CurrentThread.Name,
                    PrintController = new StandardPrintController(),
                    OriginAtMargins = false
                };

                // Señal: cuando arranca el ciclo de impresión ya quedó en cola del spooler
                printDocument.BeginPrint += (s, e) => _queuedEvt.Set();

                printDocument.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);
                printDocument.QueryPageSettings += OnQueryPageSettings;
                printDocument.PrintPage += OnPrintPage;

                printDocument.Print();

                Status = PrintJobStatus.Done;
                OnStatusUpdate("Printing Done");
            }
            catch
            {
                Status = PrintJobStatus.Failure;
                OnStatusUpdate("Printing failed");
            }
            finally
            {
                if (printDocument != null)
                {
                    printDocument.QueryPageSettings -= OnQueryPageSettings;
                    printDocument.PrintPage -= OnPrintPage;
                    printDocument.Dispose();
                }
            }
        }

        void OnPrintPage(object sender, PrintPageEventArgs e)
        {
            // Compensar hard margins del driver
            int hardX = (int)Math.Round(e.PageSettings.HardMarginX);
            int hardY = (int)Math.Round(e.PageSettings.HardMarginY);
            e.Graphics.TranslateTransform(-hardX, -hardY);

            var target = new Rectangle(0, 0, e.PageBounds.Width, e.PageBounds.Height);

            using (var bmp = new Bitmap(target.Width, target.Height, PixelFormat.Format32bppArgb))
            {
                bmp.SetResolution(100, 100);
                using (var g = Graphics.FromImage(bmp))
                {
                    g.PageUnit = GraphicsUnit.Display; // 1/100 in
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

                    Rectangle dest = new Rectangle(0, 0, target.Width, target.Height);
                    if (FitToPage && _currentFilmBox != null)
                    {
                        bool landscape = _currentFilmBox.FilmOrienation == "LANDSCAPE";
                        double pageW = target.Width, pageH = target.Height;
                        double filmRatio = landscape ? 4.0 / 3.0 : 3.0 / 4.0; // aproximado para evitar deformar feo

                        double pageRatio = pageW / pageH;
                        if (pageRatio > filmRatio)
                        {
                            int w = (int)Math.Round(pageH * filmRatio);
                            int x = (int)Math.Round((pageW - w) / 2.0);
                            dest = new Rectangle(x, 0, w, (int)pageH);
                        }
                        else
                        {
                            int h = (int)Math.Round(pageW / filmRatio);
                            int y = (int)Math.Round((pageH - h) / 2.0);
                            dest = new Rectangle(0, y, (int)pageW, h);
                        }
                    }

                    _currentFilmBox.Print(g, dest, 100);
                }

                if (_polarityReverse)
                {
                    var cm = new ColorMatrix(new float[][]
                    {
                        new float[] { -1,  0,  0, 0, 0 },
                        new float[] {  0, -1,  0, 0, 0 },
                        new float[] {  0,  0, -1, 0, 0 },
                        new float[] {  0,  0,  0, 1, 0 },
                        new float[] {  1,  1,  1, 0, 1 }
                    });
                    using (var ia = new ImageAttributes())
                    {
                        ia.SetColorMatrix(cm);
                        e.Graphics.DrawImage(bmp, target, 0, 0, bmp.Width, bmp.Height, GraphicsUnit.Pixel, ia);
                    }
                }
                else
                {
                    e.Graphics.DrawImage(bmp, target);
                }
            }

            _currentFilmBox = null;
            _currentPage++;
            e.HasMorePages = _currentPage < FilmBoxFolderList.Count;
        }

        void OnQueryPageSettings(object sender, QueryPageSettingsEventArgs e)
        {
            OnStatusUpdate(string.Format("Printing film {0} of {1}", _currentPage + 1, FilmBoxFolderList.Count));

            var filmBoxFolder = string.Format("{0}\\{1}", FullPrintJobFolder, FilmBoxFolderList[_currentPage]);
            var filmSession = FilmSession.Load(string.Format("{0}\\FilmSession.dcm", filmBoxFolder));
            _currentFilmBox = FilmBox.Load(filmSession, filmBoxFolder);

            // Leer atributos del FilmBox
            string filmSizeId = TryGetFilmSizeId(_currentFilmBox);   // ej: 10INX12IN, 24CMX30CM
            string polarity = TryGetPolarity(_currentFilmBox);     // NORMAL | REVERSE
            _polarityReverse = string.Equals(polarity, "REVERSE", StringComparison.OrdinalIgnoreCase);

            // Márgenes 0 + orientación
            e.PageSettings.Margins.Left = 0;
            e.PageSettings.Margins.Right = 0;
            e.PageSettings.Margins.Top = 0;
            e.PageSettings.Margins.Bottom = 0;
            e.PageSettings.Landscape = _currentFilmBox.FilmOrienation == "LANDSCAPE";

            // ===== Overrides por ruta =====
            if (!string.IsNullOrWhiteSpace(ForcedPaperSize))
            {
                if (TryPickPaperSize(e.PageSettings.PrinterSettings, ForcedPaperSize, out var ps))
                    e.PageSettings.PaperSize = ps;
                else if (TryParseFilmSizeId(ForcedPaperSize, out var w, out var h))
                    e.PageSettings.PaperSize = new PaperSize($"FORCED_{ForcedPaperSize}", w, h);
                else
                    Log.Warn($"No pude resolver forcePaperSize='{ForcedPaperSize}'.");
            }
            else if (!string.IsNullOrWhiteSpace(filmSizeId))
            {
                // Solo si NO hay override
                if (TryMakePaperSize(e.PageSettings.PrinterSettings, filmSizeId, out var ps))
                    e.PageSettings.PaperSize = ps;
            }

            if (!string.IsNullOrWhiteSpace(ForcedPaperSource))
            {
                if (TryPickPaperSource(e.PageSettings.PrinterSettings, ForcedPaperSource, out var src))
                    e.PageSettings.PaperSource = src;
                else
                    Log.Warn($"No encontré bandeja que matchee '{ForcedPaperSource}'.");
            }
        }

        private void DeletePrintFolder()
        {
            var folderInfo = new System.IO.DirectoryInfo(FullPrintJobFolder);
            if (folderInfo.Exists) folderInfo.Delete(true);
        }

        protected virtual void OnStatusUpdate(string info)
        {
            ExecutionStatus = Status.ToString();
            ExecutionStatusInfo = info;

            if (Status != PrintJobStatus.Failure)
                Log.Info("Print Job {0} Status {1}: {2}", SOPInstanceUID.UID.Split('.').Last(), Status, info);
            else
                Log.Error("Print Job {0} Status {1}: {2}", SOPInstanceUID.UID.Split('.').Last(), Status, info);

            StatusUpdate?.Invoke(this, new StatusUpdateEventArgs((ushort)Status, info, FilmSessionLabel, PrinterName));
        }

        // ===================== Helpers =====================

        private static string TryGetFilmSizeId(FilmBox fb)
        {
            var t = fb.GetType();
            var prop = t.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                        .FirstOrDefault(p => string.Equals(p.Name, "FilmSizeID", StringComparison.OrdinalIgnoreCase)
                                          || string.Equals(p.Name, "FilmSizeId", StringComparison.OrdinalIgnoreCase)
                                          || string.Equals(p.Name, "FilmSize", StringComparison.OrdinalIgnoreCase));
            if (prop != null)
            {
                var v = prop.GetValue(fb);
                if (v != null) return v.ToString();
            }

            var dsProp = t.GetProperty("Dataset", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
            if (dsProp?.GetValue(fb) is DicomDataset ds)
            {
                if (ds.Contains(DicomTag.FilmSizeID))
                {
                    var id = ds.Get<string>(DicomTag.FilmSizeID, null);
                    if (!string.IsNullOrEmpty(id)) return id;
                }
            }
            return null;
        }

        private static string TryGetPolarity(FilmBox fb)
        {
            var t = fb.GetType();
            var prop = t.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                        .FirstOrDefault(p => string.Equals(p.Name, "Polarity", StringComparison.OrdinalIgnoreCase));
            if (prop != null)
            {
                var v = prop.GetValue(fb);
                if (v != null) return v.ToString();
            }

            var dsProp = t.GetProperty("Dataset", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
            if (dsProp?.GetValue(fb) is DicomDataset ds)
            {
                if (ds.Contains(DicomTag.Polarity))
                {
                    var s = ds.Get<string>(DicomTag.Polarity, "NORMAL");
                    if (!string.IsNullOrWhiteSpace(s)) return s;
                }
            }
            return "NORMAL";
        }

        private static bool TryPickPaperSource(PrinterSettings ps, string wanted, out PaperSource src)
        {
            src = null;
            if (ps?.PaperSources == null) return false;
            var w = wanted.Trim().ToLowerInvariant();
            foreach (PaperSource s in ps.PaperSources)
            {
                var name = (s.SourceName ?? "").Trim().ToLowerInvariant();
                if (name.Equals(w) || name.Contains(w)) { src = s; return true; }
            }
            return false;
        }

        private static bool TryPickPaperSize(PrinterSettings ps, string wanted, out PaperSize paper)
        {
            paper = null;
            if (ps?.PaperSizes == null) return false;
            var w = wanted.Trim().ToUpperInvariant();
            foreach (PaperSize p in ps.PaperSizes)
            {
                var n = (p.PaperName ?? "").Trim().ToUpperInvariant();
                if (n.Equals(w) || n.Contains(w)) { paper = p; return true; }
                if (w == "A4" && p.Kind == PaperKind.A4) { paper = p; return true; }
                if ((w == "LETTER" || w == "LTR") && p.Kind == PaperKind.Letter) { paper = p; return true; }
                if (w == "LEGAL" && p.Kind == PaperKind.Legal) { paper = p; return true; }
                if (w == "A3" && p.Kind == PaperKind.A3) { paper = p; return true; }
            }
            return false;
        }

        private static bool TryMakePaperSize(PrinterSettings ps, string filmSizeId, out PaperSize paperSize)
        {
            paperSize = null;
            if (!TryParseFilmSizeId(filmSizeId, out var wHundInch, out var hHundInch)) return false;

            PaperSize best = null;
            int bestDelta = int.MaxValue;
            foreach (PaperSize p in ps.PaperSizes)
            {
                int dw = Math.Abs(p.Width - wHundInch);
                int dh = Math.Abs(p.Height - hHundInch);
                int delta = Math.Min(dw + dh, Math.Abs(p.Width - hHundInch) + Math.Abs(p.Height - wHundInch));
                if (delta < bestDelta) { bestDelta = delta; best = p; }
            }
            if (best != null && bestDelta <= 6) { paperSize = best; return true; }
            paperSize = new PaperSize($"DICOM_{filmSizeId}", wHundInch, hHundInch);
            return true;
        }

        private static bool TryParseFilmSizeId(string s, out int widthHundredthsInch, out int heightHundredthsInch)
        {
            widthHundredthsInch = 0; heightHundredthsInch = 0;
            if (string.IsNullOrWhiteSpace(s)) return false;
            s = s.Trim().ToUpperInvariant().Replace(" ", "");
            bool isIn = s.Contains("INX"), isCm = s.Contains("CMX");
            if (!(isIn || isCm)) return false;
            try
            {
                if (isIn)
                {
                    var parts = s.Split(new[] { "INX" }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length != 2 || !parts[1].EndsWith("IN")) return false;
                    double wIn = double.Parse(parts[0], System.Globalization.CultureInfo.InvariantCulture);
                    double hIn = double.Parse(parts[1].Replace("IN", ""), System.Globalization.CultureInfo.InvariantCulture);
                    widthHundredthsInch = (int)Math.Round(wIn * 100.0);
                    heightHundredthsInch = (int)Math.Round(hIn * 100.0);
                    return true;
                }
                else
                {
                    var parts = s.Split(new[] { "CMX" }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length != 2 || !parts[1].EndsWith("CM")) return false;
                    double wCm = double.Parse(parts[0], System.Globalization.CultureInfo.InvariantCulture);
                    double hCm = double.Parse(parts[1].Replace("CM", ""), System.Globalization.CultureInfo.InvariantCulture);
                    double wIn = wCm / 2.54, hIn = hCm / 2.54;
                    widthHundredthsInch = (int)Math.Round(wIn * 100.0);
                    heightHundredthsInch = (int)Math.Round(hIn * 100.0);
                    return true;
                }
            }
            catch { return false; }
        }
    }
}
