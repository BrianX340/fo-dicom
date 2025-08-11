using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Text.Json;
using System.Deployment.Application;
using Dicom;
using Dicom.Imaging;
using Dicom.Network;

namespace Dicom.Printing
{
    // ===================== Routing config =====================
    public class RouteItem
    {
        public string caller { get; set; }
        public string called { get; set; }
        public string printerName { get; set; }
        public string duplex { get; set; }            // "LongEdge" | "ShortEdge" | "Simplex"
        public string forcePaperSize { get; set; }    // "A4" | "Letter" | "10INX12IN" | "24CMX30CM"
        public string forceTray { get; set; }         // nombre/alias de bandeja (PaperSource.SourceName)
        public bool? fitToPage { get; set; }         // default true
        public bool? sendEventReports { get; set; }  // default false
    }

    public class PrintRoutingConfig
    {
        public List<RouteItem> routes { get; set; } = new List<RouteItem>();
    }

    internal static class RouteResolver
    {
        private static readonly object _lock = new object();
        private static PrintRoutingConfig _cache;
        private static DateTime _cacheMtimeUtc = DateTime.MinValue;

        public static string RoutesPath = ResolveRoutesPath();

        private static string ResolveRoutesPath()
        {
            try
            {
                if (ApplicationDeployment.IsNetworkDeployed)
                {
                    var act = ApplicationDeployment.CurrentDeployment.ActivationUri;
                    if (act != null && act.IsFile)
                    {
                        var root = Path.GetDirectoryName(act.LocalPath);
                        var p1 = Path.Combine(root ?? "", "routes.json");
                        if (File.Exists(p1)) { Console.WriteLine($"[Routes] Using publisher root: {p1}"); return p1; }
                    }

                    var upd = ApplicationDeployment.CurrentDeployment.UpdateLocation;
                    if (upd != null && upd.IsFile)
                    {
                        var root = Path.GetDirectoryName(upd.LocalPath);
                        var p2 = Path.Combine(root ?? "", "routes.json");
                        if (File.Exists(p2)) { Console.WriteLine($"[Routes] Using update root: {p2}"); return p2; }
                    }
                }
            }
            catch { /* ignore */ }

            var appBase = Path.Combine(AppContext.BaseDirectory, "routes.json");
            if (File.Exists(appBase)) { Console.WriteLine($"[Routes] Using AppBase: {appBase}"); return appBase; }

            try
            {
                if (ApplicationDeployment.IsNetworkDeployed)
                {
                    var dataDir = ApplicationDeployment.CurrentDeployment.DataDirectory;
                    var dataPath = Path.Combine(dataDir, "routes.json");
                    if (File.Exists(dataPath)) { Console.WriteLine($"[Routes] Using DataDirectory: {dataPath}"); return dataPath; }
                }
            }
            catch { /* ignore */ }

            Console.WriteLine($"[Routes] NOT FOUND; fallback to AppBase: {appBase}");
            return appBase;
        }

        public static RouteItem Find(string caller, string called)
        {
            var cfg = Load();
            return cfg.routes.FirstOrDefault(r =>
                string.Equals(r.caller ?? "", caller ?? "", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(r.called ?? "", called ?? "", StringComparison.OrdinalIgnoreCase));
        }

        private static PrintRoutingConfig Load()
        {
            lock (_lock)
            {
                try
                {
                    var fi = new FileInfo(RoutesPath);
                    if (!fi.Exists)
                    {
                        Console.WriteLine($"[Routes] NOT FOUND at: {RoutesPath}");
                        return _cache ?? new PrintRoutingConfig();
                    }

                    if (_cache == null || fi.LastWriteTimeUtc > _cacheMtimeUtc)
                    {
                        var json = File.ReadAllText(RoutesPath);
                        _cache = JsonSerializer.Deserialize<PrintRoutingConfig>(json, new JsonSerializerOptions
                        {
                            PropertyNameCaseInsensitive = true
                        }) ?? new PrintRoutingConfig();

                        _cacheMtimeUtc = fi.LastWriteTimeUtc;
                        Console.WriteLine($"[Routes] Loaded from: {RoutesPath}");
                    }

                    return _cache;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[Routes] Error reading '{RoutesPath}': {ex.Message}");
                    return _cache ?? new PrintRoutingConfig();
                }
            }
        }
    }

    // ===================== PrintService (SCP) =====================
    public class PrintService : DicomService, IDicomServiceProvider, IDicomNServiceProvider, IDicomCEchoProvider
    {
        public static readonly DicomTransferSyntax[] AcceptedTransferSyntaxes = new[]
        {
            DicomTransferSyntax.ExplicitVRLittleEndian,
            DicomTransferSyntax.ExplicitVRBigEndian,
            DicomTransferSyntax.ImplicitVRLittleEndian
        };

        private static DicomServer<PrintService> _server;
        public static Printer Printer { get; private set; }

        public string CallingAE { get; protected set; }
        public string CalledAE { get; protected set; }
        public System.Net.IPAddress RemoteIP { get; private set; }

        private FilmSession _filmSession;
        private readonly Dictionary<string, PrintJob> _printJobList = new Dictionary<string, PrintJob>();
        private bool _sendEventReports = false;
        private readonly object _synchRoot = new object();

        public PrintService(Stream stream, Dicom.Log.Logger log) : base(stream, log)
        {
            var pi = stream.GetType().GetProperty("Socket", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            if (pi != null)
            {
                var endPoint = ((System.Net.Sockets.Socket)pi.GetValue(stream, null)).RemoteEndPoint as System.Net.IPEndPoint;
                RemoteIP = endPoint.Address;
            }
            else
            {
                RemoteIP = new System.Net.IPAddress(new byte[] { 127, 0, 0, 1 });
            }
        }

        public static void Start(int port, string aet)
        {
            Printer = new Printer(aet);
            _server = new DicomServer<PrintService>(port);
        }

        public static void Stop()
        {
            _server?.Dispose();
        }

        // ---- Association lifecycle ----
        public void OnReceiveAssociationRequest(DicomAssociation association)
        {
            Logger.Info("Received association request from AE: {0} with IP: {1}", association.CallingAE, RemoteIP);

            // Aceptamos cualquier CalledAE y lo usamos para ruteo
            CallingAE = association.CallingAE;
            CalledAE = association.CalledAE;

            foreach (var pc in association.PresentationContexts)
            {
                if (pc.AbstractSyntax == DicomUID.Verification ||
                    pc.AbstractSyntax == DicomUID.BasicGrayscalePrintManagementMetaSOPClass ||
                    pc.AbstractSyntax == DicomUID.BasicColorPrintManagementMetaSOPClass ||
                    pc.AbstractSyntax == DicomUID.PrinterSOPClass ||
                    pc.AbstractSyntax == DicomUID.BasicFilmSessionSOPClass ||
                    pc.AbstractSyntax == DicomUID.BasicFilmBoxSOPClass ||
                    pc.AbstractSyntax == DicomUID.BasicGrayscaleImageBoxSOPClass ||
                    pc.AbstractSyntax == DicomUID.BasicColorImageBoxSOPClass)
                {
                    pc.AcceptTransferSyntaxes(AcceptedTransferSyntaxes);
                }
                else if (pc.AbstractSyntax == DicomUID.PrintJobSOPClass)
                {
                    pc.AcceptTransferSyntaxes(AcceptedTransferSyntaxes);
                    _sendEventReports = true;
                }
                else
                {
                    Logger.Warn("Requested abstract syntax {0} from {1} not supported", pc.AbstractSyntax, association.CallingAE);
                    pc.SetResult(DicomPresentationContextResult.RejectAbstractSyntaxNotSupported);
                }
            }

            Logger.Info("Accepted association request from {0} (CalledAE={1})", association.CallingAE, association.CalledAE);
            SendAssociationAccept(association);
        }

        public void OnReceiveAssociationReleaseRequest()
        {
            Clean();
            SendAssociationReleaseResponse();
        }

        public void OnReceiveAbort(DicomAbortSource source, DicomAbortReason reason)
        {
            Logger.Error("Received abort from {0}, reason is {1}", source, reason);
        }

        public void OnConnectionClosed(int errorCode)
        {
            Clean();
        }

        // ---- C-ECHO ----
        public DicomCEchoResponse OnCEchoRequest(DicomCEchoRequest request)
        {
            Logger.Info("Received verification request from AE {0} with IP: {1}", CallingAE, RemoteIP);
            return new DicomCEchoResponse(request, DicomStatus.Success);
        }

        // ---- N-CREATE ----
        public DicomNCreateResponse OnNCreateRequest(DicomNCreateRequest request)
        {
            lock (_synchRoot)
            {
                if (request.SOPClassUID == DicomUID.BasicFilmSessionSOPClass)
                {
                    return CreateFilmSession(request);
                }
                else if (request.SOPClassUID == DicomUID.BasicFilmBoxSOPClass)
                {
                    return CreateFilmBox(request);
                }
                else
                {
                    return new DicomNCreateResponse(request, DicomStatus.SOPClassNotSupported);
                }
            }
        }

        private DicomNCreateResponse CreateFilmSession(DicomNCreateRequest request)
        {
            if (_filmSession != null)
            {
                Logger.Error("Attempted to create new basic film session on association with {0}", CallingAE);
                SendAbort(DicomAbortSource.ServiceProvider, DicomAbortReason.NotSpecified);
                return new DicomNCreateResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            var pc = request.PresentationContext;
            bool isColor = pc != null && pc.AbstractSyntax == DicomUID.BasicColorPrintManagementMetaSOPClass;

            _filmSession = new FilmSession(request.SOPClassUID, request.SOPInstanceUID, request.Dataset, isColor);

            Logger.Info("Create new film session {0}", _filmSession.SOPInstanceUID.UID);

            var response = new DicomNCreateResponse(request, DicomStatus.Success);
            response.Command.Add(DicomTag.AffectedSOPInstanceUID, _filmSession.SOPInstanceUID);
            return response;
        }

        private DicomNCreateResponse CreateFilmBox(DicomNCreateRequest request)
        {
            if (_filmSession == null)
            {
                Logger.Error("A basic film session does not exist for this association {0}", CallingAE);
                SendAbort(DicomAbortSource.ServiceProvider, DicomAbortReason.NotSpecified);
                return new DicomNCreateResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            var filmBox = _filmSession.CreateFilmBox(request.SOPInstanceUID, request.Dataset);

            if (!filmBox.Initialize())
            {
                Logger.Error("Failed to initialize requested film box {0}", filmBox.SOPInstanceUID.UID);
                SendAbort(DicomAbortSource.ServiceProvider, DicomAbortReason.NotSpecified);
                return new DicomNCreateResponse(request, DicomStatus.ProcessingFailure);
            }

            Logger.Info("Created new film box {0}", filmBox.SOPInstanceUID.UID);

            var response = new DicomNCreateResponse(request, DicomStatus.Success);
            response.Command.Add(DicomTag.AffectedSOPInstanceUID, filmBox.SOPInstanceUID);
            response.Dataset = filmBox;
            return response;
        }

        // ---- N-DELETE ----
        public DicomNDeleteResponse OnNDeleteRequest(DicomNDeleteRequest request)
        {
            lock (_synchRoot)
            {
                if (request.SOPClassUID == DicomUID.BasicFilmSessionSOPClass)
                {
                    return DeleteFilmSession(request);
                }
                else if (request.SOPClassUID == DicomUID.BasicFilmBoxSOPClass)
                {
                    return DeleteFilmBox(request);
                }
                else
                {
                    return new DicomNDeleteResponse(request, DicomStatus.NoSuchSOPClass);
                }
            }
        }

        private DicomNDeleteResponse DeleteFilmBox(DicomNDeleteRequest request)
        {
            if (_filmSession == null)
            {
                Logger.Error("Can't delete a basic film session that does not exist for this association {0}", CallingAE);
                return new DicomNDeleteResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            var status = _filmSession.DeleteFilmBox(request.SOPInstanceUID)
                ? DicomStatus.Success
                : DicomStatus.NoSuchObjectInstance;

            var response = new DicomNDeleteResponse(request, status);
            response.Command.Add(DicomTag.AffectedSOPInstanceUID, request.SOPInstanceUID);
            return response;
        }

        private DicomNDeleteResponse DeleteFilmSession(DicomNDeleteRequest request)
        {
            if (_filmSession == null)
            {
                Logger.Error("Can't delete a basic film session that does not exist for this association {0}", CallingAE);
                return new DicomNDeleteResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            if (!request.SOPInstanceUID.Equals(_filmSession.SOPInstanceUID))
            {
                Logger.Error("Can't delete film session with instance UID {0} for this association {1}",
                    request.SOPInstanceUID.UID, CallingAE);
                return new DicomNDeleteResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            _filmSession = null;
            return new DicomNDeleteResponse(request, DicomStatus.Success);
        }

        // ---- N-SET ----
        public DicomNSetResponse OnNSetRequest(DicomNSetRequest request)
        {
            lock (_synchRoot)
            {
                if (request.SOPClassUID == DicomUID.BasicFilmSessionSOPClass)
                {
                    return SetFilmSession(request);
                }
                else if (request.SOPClassUID == DicomUID.BasicFilmBoxSOPClass)
                {
                    return SetFilmBox(request);
                }
                else if (request.SOPClassUID == DicomUID.BasicColorImageBoxSOPClass ||
                         request.SOPClassUID == DicomUID.BasicGrayscaleImageBoxSOPClass)
                {
                    return SetImageBox(request);
                }
                else
                {
                    return new DicomNSetResponse(request, DicomStatus.SOPClassNotSupported);
                }
            }
        }

        private DicomNSetResponse SetImageBox(DicomNSetRequest request)
        {
            if (_filmSession == null)
            {
                Logger.Error("A basic film session does not exist for this association {0}", CallingAE);
                return new DicomNSetResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            Logger.Info("Set image box {0}", request.SOPInstanceUID.UID);

            var imageBox = _filmSession.FindImageBox(request.SOPInstanceUID);
            if (imageBox == null)
            {
                Logger.Error("Received N-SET request for invalid image box instance {0} for this association {1}", request.SOPInstanceUID.UID, CallingAE);
                return new DicomNSetResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            request.Dataset.CopyTo(imageBox);
            return new DicomNSetResponse(request, DicomStatus.Success);
        }

        private DicomNSetResponse SetFilmBox(DicomNSetRequest request)
        {
            if (_filmSession == null)
            {
                Logger.Error("A basic film session does not exist for this association {0}", CallingAE);
                return new DicomNSetResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            Logger.Info("Set film box {0}", request.SOPInstanceUID.UID);
            var filmBox = _filmSession.FindFilmBox(request.SOPInstanceUID);

            if (filmBox == null)
            {
                Logger.Error("Received N-SET request for invalid film box {0} from {1}", request.SOPInstanceUID.UID, CallingAE);
                return new DicomNSetResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            request.Dataset.CopyTo(filmBox);
            filmBox.Initialize();

            var response = new DicomNSetResponse(request, DicomStatus.Success);
            response.Command.Add(DicomTag.AffectedSOPInstanceUID, filmBox.SOPInstanceUID);
            response.Command.Add(DicomTag.CommandDataSetType, (ushort)0x0202);
            response.Dataset = filmBox;
            return response;
        }

        private DicomNSetResponse SetFilmSession(DicomNSetRequest request)
        {
            if (_filmSession == null || _filmSession.SOPInstanceUID.UID != request.SOPInstanceUID.UID)
            {
                Logger.Error("A basic film session does not exist for this association {0}", CallingAE);
                return new DicomNSetResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            Logger.Info("Set film session {0}", request.SOPInstanceUID.UID);
            request.Dataset.CopyTo(_filmSession);

            return new DicomNSetResponse(request, DicomStatus.Success);
        }

        // ---- N-GET ----
        public DicomNGetResponse OnNGetRequest(DicomNGetRequest request)
        {
            lock (_synchRoot)
            {
                Logger.Info(request.ToString(true));

                if (request.SOPClassUID == DicomUID.PrinterSOPClass && request.SOPInstanceUID == DicomUID.PrinterSOPInstance)
                {
                    return GetPrinter(request);
                }
                else if (request.SOPClassUID == DicomUID.PrintJobSOPClass)
                {
                    return GetPrintJob(request);
                }
                else if (request.SOPClassUID == DicomUID.PrinterConfigurationRetrievalSOPClass && request.SOPInstanceUID == DicomUID.PrinterConfigurationRetrievalSOPInstance)
                {
                    return GetPrinterConfiguration(request);
                }
                else
                {
                    return new DicomNGetResponse(request, DicomStatus.NoSuchSOPClass);
                }
            }
        }

        private DicomNGetResponse GetPrinter(DicomNGetRequest request)
        {
            var ds = new DicomDataset();

            if (request.Attributes != null && request.Attributes.Length > 0)
            {
                foreach (var item in request.Attributes)
                {
                    var value = Printer.Get(item, "");
                    ds.Add(item, value);
                }
            }
            if (ds.Count() == 0)
            {
                ds.Add(DicomTag.PrinterStatus, Printer.PrinterStatus);
                ds.Add(DicomTag.PrinterStatusInfo, "");
                ds.Add(DicomTag.PrinterName, Printer.PrinterName);
                ds.Add(DicomTag.Manufacturer, Printer.Manufacturer);
                ds.Add(DicomTag.DateOfLastCalibration, Printer.DateTimeOfLastCalibration.Date);
                ds.Add(DicomTag.TimeOfLastCalibration, Printer.DateTimeOfLastCalibration);
                ds.Add(DicomTag.ManufacturerModelName, Printer.ManufacturerModelName);
                ds.Add(DicomTag.DeviceSerialNumber, Printer.DeviceSerialNumber);
                ds.Add(DicomTag.SoftwareVersions, Printer.SoftwareVersions);
            }

            var response = new DicomNGetResponse(request, DicomStatus.Success) { Dataset = ds };
            Logger.Info(response.ToString(true));
            return response;
        }

        private DicomNGetResponse GetPrinterConfiguration(DicomNGetRequest request)
        {
            var dataset = new DicomDataset();
            var config = new DicomDataset();
            var sequence = new DicomSequence(DicomTag.PrinterConfigurationSequence, config);
            dataset.Add(sequence);

            var response = new DicomNGetResponse(request, DicomStatus.Success);
            response.Command.Add(DicomTag.AffectedSOPInstanceUID, request.SOPInstanceUID);
            response.Dataset = dataset;

            return response;
        }

        private DicomNGetResponse GetPrintJob(DicomNGetRequest request)
        {
            if (_printJobList.ContainsKey(request.SOPInstanceUID.UID))
            {
                var printJob = _printJobList[request.SOPInstanceUID.UID];
                var dataset = new DicomDataset();

                if (request.Attributes != null && request.Attributes.Length > 0)
                {
                    foreach (var item in request.Attributes)
                    {
                        var value = printJob.Get(item, "");
                        dataset.Add(item, value);
                    }
                }

                return new DicomNGetResponse(request, DicomStatus.Success) { Dataset = dataset };
            }
            else
            {
                return new DicomNGetResponse(request, DicomStatus.NoSuchObjectInstance);
            }
        }

        // ---- N-ACTION (Print) ----
        public DicomNActionResponse OnNActionRequest(DicomNActionRequest request)
        {
            Console.WriteLine(request.ToString());

            // 1) Resolver ruta por AE
            var caller = string.IsNullOrWhiteSpace(CallingAE) ? "<unknown>" : CallingAE;
            var called = string.IsNullOrWhiteSpace(CalledAE) ? (Printer?.PrinterAet ?? "<unknown>") : CalledAE;

            Logger.Info($"Routing decision by AE: CallingAE='{caller}'  CalledAE='{called}'");

            var route = RouteResolver.Find(caller, called);
            if (route == null)
            {
                var msg = $"No route found in '{RouteResolver.RoutesPath}' for {caller}>{called}";
                Logger.Error(msg);
                return new DicomNActionResponse(request, DicomStatus.ProcessingFailure);
            }

            // 2) Aplicar impresora de Windows
            Printer.PrinterName = route.printerName;
            try
            {
                bool found = false;
                foreach (string p in PrinterSettings.InstalledPrinters)
                    if (string.Equals(p, route.printerName, StringComparison.OrdinalIgnoreCase)) { found = true; break; }
                if (!found)
                {
                    Logger.Error($"Target Windows printer not found: '{route.printerName}'.");
                    return new DicomNActionResponse(request, DicomStatus.ProcessingFailure);
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Validation printer error: {ex.Message}");
                return new DicomNActionResponse(request, DicomStatus.ProcessingFailure);
            }

            var duplexEnum = MapDuplexEnum(route.duplex);
            Logger.Info($"Route matched. Using PrinterName='{Printer.PrinterName}', Duplex='{route.duplex ?? "default"}', ForcePaper='{route.forcePaperSize ?? "-"}', ForceTray='{route.forceTray ?? "-"}'");

            if (_filmSession == null)
            {
                Logger.Error("A basic film session does not exist for this association {0}", CallingAE);
                return new DicomNActionResponse(request, DicomStatus.InvalidObjectInstance);
            }

            lock (_synchRoot)
            {
                try
                {
                    // 3) Recolectar FilmBoxes a imprimir
                    var filmBoxList = new List<FilmBox>();
                    if (request.SOPClassUID == DicomUID.BasicFilmSessionSOPClass && request.ActionTypeID == 0x0001)
                    {
                        Logger.Info("Creating new print job for film session {0}", _filmSession.SOPInstanceUID.UID);
                        filmBoxList.AddRange(_filmSession.BasicFilmBoxes);
                    }
                    else if (request.SOPClassUID == DicomUID.BasicFilmBoxSOPClass && request.ActionTypeID == 0x0001)
                    {
                        Logger.Info("Creating new print job for film box {0}", request.SOPInstanceUID.UID);
                        var filmBox = _filmSession.FindFilmBox(request.SOPInstanceUID);
                        if (filmBox != null) filmBoxList.Add(filmBox);
                        else
                        {
                            Logger.Error("Received N-ACTION request for invalid film box {0} from {1}", request.SOPInstanceUID.UID, CallingAE);
                            return new DicomNActionResponse(request, DicomStatus.NoSuchObjectInstance);
                        }
                    }
                    else
                    {
                        if (request.ActionTypeID != 0x0001)
                        {
                            Logger.Error("Received N-ACTION request for invalid action type {0} from {1}", request.ActionTypeID, CallingAE);
                            return new DicomNActionResponse(request, DicomStatus.NoSuchActionType);
                        }
                        else
                        {
                            Logger.Error("Received N-ACTION request for invalid SOP class {0} from {1}", request.SOPClassUID, CallingAE);
                            return new DicomNActionResponse(request, DicomStatus.NoSuchSOPClass);
                        }
                    }

                    // 4) (Opcional) export de preview a JPEG
                    try
                    {
                        var outDir = Path.Combine(AppContext.BaseDirectory, "prints",
                            DateTime.Now.ToString("yyyyMMdd_HHmmss") + $"_{caller}_{called}");
                        Directory.CreateDirectory(outDir);

                        int saved = TrySaveFilmBoxesJpeg_SystemDrawing(filmBoxList, outDir, Logger);
                        Logger.Info($"Preview export: saved {saved} JPEG file(s) to {outDir}");
                    }
                    catch (Exception exPrev)
                    {
                        Logger.Warn($"Preview export failed (non-fatal): {exPrev.Message}");
                    }

                    // 5) Lanzar impresión (y responder Success apenas entra en cola)
                    var printJob = new PrintJob(null, Printer, CallingAE, Logger)
                    {
                        SendNEventReport = route.sendEventReports ?? false,
                        Duplex = duplexEnum,
                        ForcedPaperSize = route.forcePaperSize,
                        ForcedPaperSource = route.forceTray,
                        FitToPage = route.fitToPage ?? true
                    };
                    printJob.StatusUpdate += OnPrintJobStatusUpdate;
                    _printJobList[printJob.SOPInstanceUID.UID] = printJob;

                    printJob.Print(filmBoxList);

                    // Espera breve a que el job entre al spooler
                    printJob.WaitUntilQueued(TimeSpan.FromSeconds(3));

                    // 6) Responder OK (Success) sin esperar impresión física
                    var result = new DicomDataset();
                    result.Add(new DicomSequence(new DicomTag(0x2100, 0x0500),
                        new DicomDataset(new DicomUniqueIdentifier(DicomTag.ReferencedSOPClassUID, DicomUID.PrintJobSOPClass)),
                        new DicomDataset(new DicomUniqueIdentifier(DicomTag.ReferencedSOPInstanceUID, printJob.SOPInstanceUID))));

                    var response = new DicomNActionResponse(request, DicomStatus.Success);
                    response.Command.Add(DicomTag.AffectedSOPInstanceUID, printJob.SOPInstanceUID);
                    response.Dataset = result;
                    return response;
                }
                catch (Exception ex)
                {
                    Logger.Error("Error occured during N-ACTION {0} for SOP class {1} and instance {2}",
                        request.ActionTypeID, request.SOPClassUID.UID, request.SOPInstanceUID.UID);
                    Logger.Error(ex.Message);
                    return new DicomNActionResponse(request, DicomStatus.ProcessingFailure);
                }
            }
        }

        private void OnPrintJobStatusUpdate(object sender, StatusUpdateEventArgs e)
        {
            var printJob = sender as PrintJob;
            if (printJob != null && printJob.SendNEventReport)
            {
                var reportRequest = new DicomNEventReportRequest(printJob.SOPClassUID, printJob.SOPInstanceUID, e.EventTypeId);
                var ds = new DicomDataset();
                ds.Add(DicomTag.ExecutionStatusInfo, e.ExecutionStatusInfo);
                ds.Add(DicomTag.FilmSessionLabel, e.FilmSessionLabel);
                ds.Add(DicomTag.PrinterName, e.PrinterName);

                reportRequest.Dataset = ds;
                SendRequest(reportRequest);
            }
        }

        public void Clean()
        {
            lock (_synchRoot)
            {
                _filmSession = null;
                _printJobList.Clear();
            }
        }

        public DicomNEventReportResponse OnNEventReportRequest(DicomNEventReportRequest request)
        {
            throw new NotImplementedException();
        }

        // ===================== Helpers =====================
        private static Duplex? MapDuplexEnum(string duplex)
        {
            if (string.IsNullOrWhiteSpace(duplex)) return null;

            switch (duplex.Trim().ToLowerInvariant())
            {
                case "longedge":
                case "twosidedlongedge":
                case "long":
                    return Duplex.Vertical;   // borde largo
                case "shortedge":
                case "twosidedshortedge":
                case "short":
                    return Duplex.Horizontal; // borde corto
                case "simplex":
                case "onesided":
                    return Duplex.Simplex;
                default:
                    return null;
            }
        }

        private static int TrySaveFilmBoxesJpeg_SystemDrawing(IEnumerable<FilmBox> filmBoxes, string outDir, Dicom.Log.Logger logger)
        {
            int saved = 0;
            int fbIndex = 0;

            foreach (var fb in filmBoxes ?? Enumerable.Empty<FilmBox>())
            {
                fbIndex++;

                IEnumerable<DicomDataset> boxes = null;

                var prop = fb.GetType().GetProperty("BasicImageBoxes") ?? fb.GetType().GetProperty("ImageBoxes");
                if (prop != null)
                {
                    var raw = prop.GetValue(fb) as System.Collections.IEnumerable;
                    if (raw != null)
                    {
                        var list = new List<DicomDataset>();
                        foreach (var o in raw)
                        {
                            if (o is DicomDataset ds) list.Add(ds);
                            else
                            {
                                var dsProp = o?.GetType().GetProperty("Dataset");
                                if (dsProp?.GetValue(o) is DicomDataset inner) list.Add(inner);
                            }
                        }
                        boxes = list;
                    }
                }

                if ((boxes == null || !boxes.Any()) && fb is DicomDataset fbDs)
                {
                    var dsFromSeq = new List<DicomDataset>();
                    foreach (var item in fbDs)
                        if (item is DicomSequence seq && seq.Items != null)
                            dsFromSeq.AddRange(seq.Items);
                    boxes = dsFromSeq;
                }

                if (boxes == null || !boxes.Any())
                {
                    logger.Warn($"FilmBox[{fbIndex}]: no image boxes found to preview.");
                    continue;
                }

                int ibIndex = 0;
                foreach (var ib in boxes)
                {
                    ibIndex++;
                    try
                    {
                        var dsWithPixels = FindFirstPixelDataDataset(ib);
                        if (dsWithPixels == null)
                        {
                            logger.Warn($"FilmBox[{fbIndex}] ImageBox[{ibIndex}]: no PixelData present (even in nested sequences).");
                            continue;
                        }

                        var dicomImage = new DicomImage(dsWithPixels);

                        using (var bmp = dicomImage.RenderImage() as Bitmap)
                        {
                            if (bmp == null)
                            {
                                logger.Warn($"FilmBox[{fbIndex}] ImageBox[{ibIndex}]: Rendered image could not be cast to Bitmap.");
                                continue;
                            }
                            var fileName = Path.Combine(outDir, $"film{fbIndex:00}_img{ibIndex:00}.jpeg");
                            bmp.Save(fileName, ImageFormat.Jpeg);
                            saved++;
                        }
                    }
                    catch (Exception exIB)
                    {
                        logger.Warn($"FilmBox[{fbIndex}] ImageBox[{ibIndex}] preview failed: {exIB.Message}");
                    }
                }
            }

            return saved;
        }

        private static DicomDataset FindFirstPixelDataDataset(DicomDataset root)
        {
            if (root == null) return null;

            if (root.Contains(DicomTag.PixelData))
                return root;

            foreach (var element in root)
            {
                if (element is DicomSequence seq && seq.Items != null)
                {
                    foreach (var item in seq.Items)
                    {
                        var found = FindFirstPixelDataDataset(item);
                        if (found != null) return found;
                    }
                }
            }
            return null;
        }
    }
}
