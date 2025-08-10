﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.Json;
using Dicom;
using Dicom.Network;
using System.Drawing;
using System.Drawing.Imaging;
using Dicom.Imaging;
using Newtonsoft.Json;
using System.Drawing.Printing;
using System.Deployment.Application;
//using FellowOakDicom.Drawing;



namespace Dicom.Printing
{
    // -------------------- Routing config models --------------------
    public class RouteItem
    {
        public string caller { get; set; }
        public string called { get; set; }
        public string printerName { get; set; }
        public string duplex { get; set; } // "LongEdge" | "ShortEdge" | "Simplex"
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
            // 1) ClickOnce: carpeta RAÍZ donde vive el .application (lo que vos querés)
            try
            {
                if (ApplicationDeployment.IsNetworkDeployed)
                {
                    var act = ApplicationDeployment.CurrentDeployment.ActivationUri;   // p.ej. file:///C:/PrintServerX340/Print%20SCP.application
                    if (act != null && act.IsFile)
                    {
                        var publisherRoot = Path.GetDirectoryName(act.LocalPath);      // C:\PrintServerX340
                        var p1 = Path.Combine(publisherRoot ?? "", "routes.json");
                        if (File.Exists(p1))
                        {
                            Console.WriteLine($"[Routes] Using publisher root: {p1}");
                            return p1;
                        }
                    }

                    // 1b) fallback: UpdateLocation (a veces ActivationUri viene null)
                    var upd = ApplicationDeployment.CurrentDeployment.UpdateLocation;  // idem, puede ser file://
                    if (upd != null && upd.IsFile)
                    {
                        var publisherRoot = Path.GetDirectoryName(upd.LocalPath);
                        var p2 = Path.Combine(publisherRoot ?? "", "routes.json");
                        if (File.Exists(p2))
                        {
                            Console.WriteLine($"[Routes] Using update root: {p2}");
                            return p2;
                        }
                    }
                }
            }
            catch { /* ignorar */ }

            // 2) AppBase (carpeta del exe en cache ClickOnce o en Debug)
            var appBase = Path.Combine(AppContext.BaseDirectory, "routes.json");
            if (File.Exists(appBase))
            {
                Console.WriteLine($"[Routes] Using AppBase: {appBase}");
                return appBase;
            }

            // 3) ClickOnce DataDirectory (por si lo publicaste como Data File)
            try
            {
                if (ApplicationDeployment.IsNetworkDeployed)
                {
                    var dataDir = ApplicationDeployment.CurrentDeployment.DataDirectory;
                    var dataPath = Path.Combine(dataDir, "routes.json");
                    if (File.Exists(dataPath))
                    {
                        Console.WriteLine($"[Routes] Using DataDirectory: {dataPath}");
                        return dataPath;
                    }
                }
            }
            catch { /* ignorar */ }

            // 4) Último recurso: devolvemos AppBase aunque no exista (para que el log diga dónde buscó)
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
                        _cache = System.Text.Json.JsonSerializer.Deserialize<PrintRoutingConfig>(json, new JsonSerializerOptions
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





    public class PrintService : DicomService, IDicomServiceProvider, IDicomNServiceProvider, IDicomCEchoProvider
    {
        #region Properties and Attributes
        public static readonly DicomTransferSyntax[] AcceptedTransferSyntaxes = new DicomTransferSyntax[]
		{
			DicomTransferSyntax.ExplicitVRLittleEndian,
			DicomTransferSyntax.ExplicitVRBigEndian,
			DicomTransferSyntax.ImplicitVRLittleEndian
		};
        public static readonly DicomTransferSyntax[] AcceptedImageTransferSyntaxes = new DicomTransferSyntax[]
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

        private Dictionary<string, PrintJob> _printJobList = new Dictionary<string, PrintJob>();

        private bool _sendEventReports = false;
        private readonly object _synchRoot = new object();
        #endregion

        #region Constructors and Initialization

        public PrintService(System.IO.Stream stream, Dicom.Log.Logger log)
            : base(stream, log)
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
            _server.Dispose();
        }

        #endregion

        #region IDicomServiceProvider Members

        public void OnReceiveAssociationRequest(DicomAssociation association)
        {
            this.Logger.Info("Received association request from AE: {0} with IP: {1} ", association.CallingAE, RemoteIP);

            if (Printer.PrinterAet != association.CalledAE)
            {
                this.Logger.Error("Association with {0} rejected since requested printer {1} not found",
                    association.CallingAE, association.CalledAE);
                SendAssociationReject(DicomRejectResult.Permanent, DicomRejectSource.ServiceUser, DicomRejectReason.CalledAENotRecognized);
                return;
            }

            CallingAE = association.CallingAE;
            CalledAE = Printer.PrinterAet;

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
                    this.Logger.Warn("Requested abstract syntax {0} from {1} not supported", pc.AbstractSyntax, association.CallingAE);
                    pc.SetResult(DicomPresentationContextResult.RejectAbstractSyntaxNotSupported);
                }
            }

            this.Logger.Info("Accepted association request from {0}", association.CallingAE);
            SendAssociationAccept(association);
        }

        public void OnReceiveAssociationReleaseRequest()
        {
            Clean();
            SendAssociationReleaseResponse();
        }

        public void OnReceiveAbort(DicomAbortSource source, DicomAbortReason reason)
        {
            this.Logger.Error("Received abort from {0}, reason is {1}", source, reason);
        }

        public void OnConnectionClosed(int errorCode)
        {
            Clean();
        }

        #endregion

        #region IDicomCEchoProvider Members

        public DicomCEchoResponse OnCEchoRequest(DicomCEchoRequest request)
        {
            this.Logger.Info("Received verification request from AE {0} with IP: {1}", CallingAE, RemoteIP);
            return new DicomCEchoResponse(request, DicomStatus.Success);
        }

        #endregion

        #region N-CREATE requests handlers

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
                this.Logger.Error("Attemted to create new basic film session on association with {0}", CallingAE);
                SendAbort(DicomAbortSource.ServiceProvider, DicomAbortReason.NotSpecified);
                return new DicomNCreateResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            var pc = request.PresentationContext;
            bool isColor = pc != null && pc.AbstractSyntax == DicomUID.BasicColorPrintManagementMetaSOPClass;

            _filmSession = new FilmSession(request.SOPClassUID, request.SOPInstanceUID, request.Dataset, isColor);

            this.Logger.Info("Create new film session {0}", _filmSession.SOPInstanceUID.UID);

            var response = new DicomNCreateResponse(request, DicomStatus.Success);
            response.Command.Add(DicomTag.AffectedSOPInstanceUID, _filmSession.SOPInstanceUID);
            return response;
        }

        private DicomNCreateResponse CreateFilmBox(DicomNCreateRequest request)
        {
            if (_filmSession == null)
            {
                this.Logger.Error("A basic film session does not exist for this association {0}", CallingAE);
                SendAbort(DicomAbortSource.ServiceProvider, DicomAbortReason.NotSpecified);
                return new DicomNCreateResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            var filmBox = _filmSession.CreateFilmBox(request.SOPInstanceUID, request.Dataset);

            if (!filmBox.Initialize())
            {
                this.Logger.Error("Failed to initialize requested film box {0}", filmBox.SOPInstanceUID.UID);
                SendAbort(DicomAbortSource.ServiceProvider, DicomAbortReason.NotSpecified);
                return new DicomNCreateResponse(request, DicomStatus.ProcessingFailure);
            }

            this.Logger.Info("Created new film box {0}", filmBox.SOPInstanceUID.UID);

            var response = new DicomNCreateResponse(request, DicomStatus.Success);
            response.Command.Add(DicomTag.AffectedSOPInstanceUID, filmBox.SOPInstanceUID);
            response.Dataset = filmBox;
            return response;
        }

        #endregion

        #region N-DELETE request handler

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
                this.Logger.Error("Can't delete a basic film session doesnot exist for this association {0}", CallingAE);
                return new DicomNDeleteResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            DicomStatus status;
            if (_filmSession.DeleteFilmBox(request.SOPInstanceUID))
            {
                status = DicomStatus.Success;
            }
            else
            {
                status = DicomStatus.NoSuchObjectInstance;
            }
            var response = new DicomNDeleteResponse(request, status);

            response.Command.Add(DicomTag.AffectedSOPInstanceUID, request.SOPInstanceUID);
            return response;
        }

        private DicomNDeleteResponse DeleteFilmSession(DicomNDeleteRequest request)
        {
            if (_filmSession == null)
            {
                this.Logger.Error("Can't delete a basic film session doesnot exist for this association {0}", CallingAE);
                return new DicomNDeleteResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            if (!request.SOPInstanceUID.Equals(_filmSession.SOPInstanceUID))
            {
                this.Logger.Error("Can't delete a basic film session with instace UID {0} doesnot exist for this association {1}",
                    request.SOPInstanceUID.UID, CallingAE);
                return new DicomNDeleteResponse(request, DicomStatus.NoSuchObjectInstance);
            }
            _filmSession = null;

            return new DicomNDeleteResponse(request, DicomStatus.Success);
        }

        #endregion

        #region N-SET request handler

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
                this.Logger.Error("A basic film session does not exist for this association {0}", CallingAE);
                return new DicomNSetResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            this.Logger.Info("Set image box {0}", request.SOPInstanceUID.UID);

            var imageBox = _filmSession.FindImageBox(request.SOPInstanceUID);
            if (imageBox == null)
            {
                this.Logger.Error("Received N-SET request for invalid image box instance {0} for this association {1}", request.SOPInstanceUID.UID, CallingAE);
                return new DicomNSetResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            request.Dataset.CopyTo(imageBox);

            return new DicomNSetResponse(request, DicomStatus.Success);
        }

        private DicomNSetResponse SetFilmBox(DicomNSetRequest request)
        {
            if (_filmSession == null)
            {
                this.Logger.Error("A basic film session does not exist for this association {0}", CallingAE);
                return new DicomNSetResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            this.Logger.Info("Set film box {0}", request.SOPInstanceUID.UID);
            var filmBox = _filmSession.FindFilmBox(request.SOPInstanceUID);

            if (filmBox == null)
            {
                this.Logger.Error("Received N-SET request for invalid film box {0} from {1}", request.SOPInstanceUID.UID, CallingAE);
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
                this.Logger.Error("A basic film session does not exist for this association {0}", CallingAE);
                return new DicomNSetResponse(request, DicomStatus.NoSuchObjectInstance);
            }

            this.Logger.Info("Set film session {0}", request.SOPInstanceUID.UID);
            request.Dataset.CopyTo(_filmSession);

            return new DicomNSetResponse(request, DicomStatus.Success);
        }

        #endregion

        #region N-GET request handler

        public DicomNGetResponse OnNGetRequest(DicomNGetRequest request)
        {
            lock (_synchRoot)
            {
                this.Logger.Info(request.ToString(true));

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

            var sb = new System.Text.StringBuilder();
            if (request.Attributes != null && request.Attributes.Length > 0)
            {
                foreach (var item in request.Attributes)
                {
                    sb.AppendFormat("GetPrinter attribute {0} requested", item);
                    sb.AppendLine();
                    var value = Printer.Get(item, "");
                    ds.Add(item, value);
                }

                Logger.Info(sb.ToString());
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

            var response = new DicomNGetResponse(request, DicomStatus.Success);
            response.Dataset = ds;

            this.Logger.Info(response.ToString(true));
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

                var sb = new System.Text.StringBuilder();
                var dataset = new DicomDataset();

                if (request.Attributes != null && request.Attributes.Length > 0)
                {
                    foreach (var item in request.Attributes)
                    {
                        sb.AppendFormat("GetPrintJob attribute {0} requested", item);
                        sb.AppendLine();
                        var value = printJob.Get(item, "");
                        dataset.Add(item, value);
                    }

                    Logger.Info(sb.ToString());
                }

                var response = new DicomNGetResponse(request, DicomStatus.Success);
                response.Dataset = dataset;
                return response;
            }
            else
            {
                var response = new DicomNGetResponse(request, DicomStatus.NoSuchObjectInstance);
                return response;
            }
        }

        #endregion

        #region N-ACTION request handler

        private static int TrySaveFilmBoxesJpeg_SystemDrawing(IEnumerable<FilmBox> filmBoxes, string outDir, Dicom.Log.Logger logger)
        {
            int saved = 0;
            int fbIndex = 0;

            foreach (var fb in filmBoxes ?? Enumerable.Empty<FilmBox>())
            {
                fbIndex++;

                // 1) Intentar obtener los ImageBoxes desde propiedades comunes
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

                // 2) Fallback: buscar datasets dentro de secuencias del propio FilmBox
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
                        // Buscar recursivamente el primer dataset con PixelData
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

        // Busca recursivamente PixelData en el dataset y sus secuencias
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

        private static bool ValidateInstalledPrinter(string printerName, out string reason)
        {
            reason = null;
            if (string.IsNullOrWhiteSpace(printerName))
            {
                reason = "printerName is empty";
                return false;
            }

            try
            {
                foreach (string p in PrinterSettings.InstalledPrinters)
                {
                    if (string.Equals(p, printerName, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
                reason = "not found in InstalledPrinters";
                return false;
            }
            catch (Exception ex)
            {
                reason = ex.Message;
                return false;
            }
        }

        private static Duplex? MapDuplexEnum(string duplex)
        {
            if (string.IsNullOrWhiteSpace(duplex)) return null;

            switch (duplex.Trim().ToLowerInvariant())
            {
                case "longedge":
                case "twosidedlongedge":
                case "long":
                    return Duplex.Vertical;      // libro (borde largo)
                case "shortedge":
                case "twosidedshortedge":
                case "short":
                    return Duplex.Horizontal;    // bloc/calendario (borde corto)
                case "simplex":
                case "onesided":
                    return Duplex.Simplex;
                default:
                    return null;
            }
        }


        public DicomNActionResponse OnNActionRequest(DicomNActionRequest request)
        {
            Console.WriteLine(request.ToString());

            // ================================================================
            // Ruteo y perfil segun caller/called leidos desde routes.json
            var caller = string.IsNullOrWhiteSpace(CallingAE) ? "<unknown>" : CallingAE;
            var called = string.IsNullOrWhiteSpace(CalledAE) ? (Printer?.PrinterAet ?? "<unknown>") : CalledAE;

            this.Logger.Info($"Routing decision by AE: CallingAE='{caller}'  CalledAE='{called}'");

            var route = RouteResolver.Find(caller, called);
            if (route == null)
            {
                var msg = $"No route found in '{RouteResolver.RoutesPath}' for {caller}>{called}";
                this.Logger.Error(msg);
                return new DicomNActionResponse(request, DicomStatus.ProcessingFailure);
            }

            // 1) Aplicar printerName al objeto Printer (tu lógica actual por reflexión)
            if (!TryApplyPrinterName(route.printerName, out var applyErr))
            {
                this.Logger.Error($"Failed to apply printerName='{route.printerName}': {applyErr}");
                return new DicomNActionResponse(request, DicomStatus.ProcessingFailure);
            }

            // 2) Validar que la impresora exista en Windows (impresora digital)
            if (!ValidateInstalledPrinter(route.printerName, out var notFoundReason))
            {
                this.Logger.Error($"Target Windows printer not found: '{route.printerName}'. {notFoundReason}");
                return new DicomNActionResponse(request, DicomStatus.ProcessingFailure);
            }

            // 3) Intentar aplicar dúplex si tu stack lo permite
            if (!string.IsNullOrWhiteSpace(route.duplex))
            {
                if (!TryApplyDuplex(route.duplex, out var duplexErr))
                    this.Logger.Warn($"Could not apply duplex='{route.duplex}': {duplexErr}");

                // Si PrintJob o Printer aceptan un Duplex nativo (System.Drawing.Printing.Duplex) vía reflexión:
                var duplexEnum = MapDuplexEnum(route.duplex); // null si no reconocemos el texto
                if (duplexEnum.HasValue)
                {
                    try
                    {
                        // Intentar en Printer (propiedad "DuplexingMode" o similar)
                        var prop = Printer.GetType().GetProperty("DuplexingMode");
                        if (prop != null && prop.CanWrite && prop.PropertyType == typeof(Duplex))
                        {
                            prop.SetValue(Printer, duplexEnum.Value);
                        }
                        else
                        {
                            // Intentar método en Printer o PrintJob si existiera (ejemplos posibles)
                            var setMethod = Printer.GetType().GetMethod("SetDuplex", new[] { typeof(Duplex) });
                            if (setMethod != null)
                                setMethod.Invoke(Printer, new object[] { duplexEnum.Value });
                        }
                    }
                    catch (Exception ex)
                    {
                        this.Logger.Warn($"Could not set native Duplex on Printer: {ex.Message}");
                    }
                }
            }

            this.Logger.Info($"Route matched. Using PrinterName='{Printer.PrinterName}', Duplex='{route.duplex ?? "default"}'");
            // ================================================================

            if (_filmSession == null)
            {
                this.Logger.Error("A basic film session does not exist for this association {0}", CallingAE);
                return new DicomNActionResponse(request, DicomStatus.InvalidObjectInstance);
            }

            lock (_synchRoot)
            {
                try
                {
                    var filmBoxList = new List<FilmBox>();
                    if (request.SOPClassUID == DicomUID.BasicFilmSessionSOPClass && request.ActionTypeID == 0x0001)
                    {
                        this.Logger.Info("Creating new print job for film session {0}", _filmSession.SOPInstanceUID.UID);
                        filmBoxList.AddRange(_filmSession.BasicFilmBoxes);
                    }
                    else if (request.SOPClassUID == DicomUID.BasicFilmBoxSOPClass && request.ActionTypeID == 0x0001)
                    {
                        this.Logger.Info("Creating new print job for film box {0}", request.SOPInstanceUID.UID);

                        var filmBox = _filmSession.FindFilmBox(request.SOPInstanceUID);
                        if (filmBox != null) filmBoxList.Add(filmBox);
                        else
                        {
                            this.Logger.Error("Received N-ACTION request for invalid film box {0} from {1}", request.SOPInstanceUID.UID, CallingAE);
                            return new DicomNActionResponse(request, DicomStatus.NoSuchObjectInstance);
                        }
                    }
                    else
                    {
                        if (request.ActionTypeID != 0x0001)
                        {
                            this.Logger.Error("Received N-ACTION request for invalid action type {0} from {1}", request.ActionTypeID, CallingAE);
                            return new DicomNActionResponse(request, DicomStatus.NoSuchActionType);
                        }
                        else
                        {
                            this.Logger.Error("Received N-ACTION request for invalid SOP class {0} from {1}", request.SOPClassUID, CallingAE);
                            return new DicomNActionResponse(request, DicomStatus.NoSuchSOPClass);
                        }
                    }

                    // ------------------------------------------------
                    // (Opcional) Previews a JPEG que ya comprobaste OK
                    try
                    {
                        var outDir = Path.Combine(AppContext.BaseDirectory, "prints",
                            DateTime.Now.ToString("yyyyMMdd_HHmmss") + $"_{caller}_{called}");
                        Directory.CreateDirectory(outDir);

                        int saved = TrySaveFilmBoxesJpeg_SystemDrawing(filmBoxList, outDir, this.Logger);
                        this.Logger.Info($"Preview export: saved {saved} JPEG file(s) to {outDir}");
                    }
                    catch (Exception exPrev)
                    {
                        this.Logger.Warn($"Preview export failed (non-fatal): {exPrev.Message}");
                    }
                    // ------------------------------------------------

                    // 4) Lanzar la impresión con tu PrintJob (usa Printer.PrinterName ya seteado)
                    var printJob = new PrintJob(null, Printer, CallingAE, this.Logger);
                    printJob.SendNEventReport = _sendEventReports;
                    printJob.StatusUpdate += OnPrintJobStatusUpdate;

                    // (Opcional) Intentar pasar Duplex nativo al printJob si lo soporta:
                    try
                    {
                        var duplexEnum = MapDuplexEnum(route.duplex);
                        if (duplexEnum.HasValue)
                        {
                            var pjProp = printJob.GetType().GetProperty("Duplex");
                            if (pjProp != null && pjProp.CanWrite && pjProp.PropertyType == typeof(Duplex))
                                pjProp.SetValue(printJob, duplexEnum.Value);
                        }
                    }
                    catch (Exception exDp)
                    {
                        this.Logger.Warn($"Could not set Duplex on PrintJob: {exDp.Message}");
                    }

                    printJob.Print(filmBoxList);

                    if (printJob.Error == null)
                    {
                        // 5) Responder al invocador con Success (y referenciar el Print Job SOP si corresponde)
                        var result = new DicomDataset();
                        result.Add(new DicomSequence(new DicomTag(0x2100, 0x0500),
                            new DicomDataset(new DicomUniqueIdentifier(DicomTag.ReferencedSOPClassUID, DicomUID.PrintJobSOPClass)),
                            new DicomDataset(new DicomUniqueIdentifier(DicomTag.ReferencedSOPInstanceUID, printJob.SOPInstanceUID))));

                        var response = new DicomNActionResponse(request, DicomStatus.Success);
                        response.Command.Add(DicomTag.AffectedSOPInstanceUID, printJob.SOPInstanceUID);
                        response.Dataset = result;

                        return response;
                    }
                    else
                    {
                        throw printJob.Error;
                    }
                }
                catch (Exception ex)
                {
                    this.Logger.Error("Error occured during N-ACTION {0} for SOP class {1} and instance {2}",
                        request.ActionTypeID, request.SOPClassUID.UID, request.SOPInstanceUID.UID);
                    this.Logger.Error(ex.Message);
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
                this.SendRequest(reportRequest);
            }
        }

        // Intenta aplicar el nombre de impresora a la instancia Printer
        private bool TryApplyPrinterName(string printerName, out string error)
        {
            error = null;
            if (string.IsNullOrWhiteSpace(printerName))
            {
                error = "printerName is empty";
                return false;
            }

            try
            {
                var prop = Printer.GetType().GetProperty("PrinterName");
                if (prop == null)
                {
                    error = "Printer.PrinterName property not found";
                    return false;
                }
                if (!prop.CanWrite)
                {
                    error = "Printer.PrinterName is read-only";
                    return false;
                }

                prop.SetValue(Printer, printerName);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        // Intenta aplicar duplex si tu clase Printer lo soporta. Si no, solo loggea.
        private bool TryApplyDuplex(string duplex, out string error)
        {
            error = null;
            if (string.IsNullOrWhiteSpace(duplex)) return true;

            try
            {
                // Intentar encontrar una propiedad o método relacionado a Duplex
                // Ejemplos posibles (descomentá si existen):
                // var prop = Printer.GetType().GetProperty("DuplexingMode");
                // if (prop != null && prop.CanWrite) { prop.SetValue(Printer, duplex); return true; }

                // var method = Printer.GetType().GetMethod("SetDuplex");
                // if (method != null) { method.Invoke(Printer, new object[] { duplex }); return true; }

                // Si no hay soporte explícito, no consideramos error fatal:
                this.Logger.Info($"Requested duplex='{duplex}' (no direct setter found on Printer).");
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }
        #endregion

        public void Clean()
        {
            lock (_synchRoot)
            {
                if (_filmSession != null)
                {
                    _filmSession = null;
                }
                _printJobList.Clear();
            }
        }

        #region IDicomNServiceProvider Members

        public DicomNEventReportResponse OnNEventReportRequest(DicomNEventReportRequest request)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
