Este repositorio fue modificado de un fork a https://github.com/hdesouky/fo-dicom/tree/master
---------------------------

DICOM Print SCP — Guía rápida

Qué resuelve
Permite que equipos de imagen (ecógrafos, RM, TC, etc.) impriman en impresoras de Windows.
El ruteo se hace por AE Titles: según el Calling AE (equipo) y el Called AE (nombre lógico de la “impresora DICOM”), el servidor elige qué impresora usar y con qué tamaño/bandeja.

1) Configurá routes.json
Cada ítem es una regla. Si llega un trabajo con caller (Calling AE del equipo) y called (Called AE configurado en el equipo), se envía a printerName con los overrides indicados.

Ejemplo con dos ecógrafos (AEs distintos) y dos impresoras físicas.
- Ecógrafo A → Calling AE = ECOGA
- Ecógrafo B → Calling AE = ECOGB

{
  "routes": [
    {
      "caller": "ECOGA",
      "called": "HUASSCOLORTESTA4",
      "printerName": "HUA_SS_Color",
      "duplex": "LongEdge",

      "forcePaperSize": "A4",
      "forceTray": "Bandeja 2",
      "fitToPage": true,
      "sendEventReports": false
    },
    {
      "caller": "ECOGA",
      "called": "HUASSCOLORTESTA3",
      "printerName": "HUA_SS_Color",
      "duplex": "LongEdge",

      "forcePaperSize": "A3",
      "forceTray": "",
      "fitToPage": true,
      "sendEventReports": false
    },
    {
      "caller": "ECOGB",
      "called": "HUAPBCOLORTESTA4",
      "printerName": "HUA_PB_Color",
      "duplex": "LongEdge",

      "forcePaperSize": "A4",
      "forceTray": "Bandeja 2",
      "fitToPage": true,
      "sendEventReports": false
    }
  ]
}

Campos clave (resumen):
- caller: Calling AE del equipo.
- called: Called AE que vas a configurar en el equipo para esta “impresora”.
- printerName: nombre exacto de la impresora en Windows.
- forcePaperSize: A3/A4/Letter o FilmSizeID (10INX12IN, 24CMX30CM).
- forceTray: nombre de bandeja (vacío si querés que el driver elija).
- fitToPage: ajusta manteniendo aspecto.
- duplex: LongEdge | ShortEdge | Simplex.
- sendEventReports: habilita o no N-EVENT-REPORT (usualmente false).

2) Configurá los equipos (SCU)
Servidor DICOM Print SCP
- IP: 11.11.11.11 (ejemplo)
- Puerto: 7250

Ecógrafo A
- Calling AE: ECOGA
- Agregar “impresoras” DICOM:
  - A4 → IP 11.11.11.11, Puerto 7250, Called AE HUASSCOLORTESTA4
  - A3 → IP 11.11.11.11, Puerto 7250, Called AE HUASSCOLORTESTA3

Ecógrafo B
- Calling AE: ECOGB
- Agregar:
  - A4 → IP 11.11.11.11, Puerto 7250, Called AE HUAPBCOLORTESTA4

Cada Called AE representa una configuración distinta (tamaño/bandeja/impresora), aunque apunte a la misma impresora física.

3) Probar
Enviá una impresión desde el equipo al Called AE deseado.
El servidor responde OK (Success) apenas el trabajo entra al spooler de Windows.

Tips y resolución de problemas
- A3 sale “chico” centrado → Es A4 dentro de A3.
  * Usá forcePaperSize: "A3" y dejá forceTray vacío o una bandeja que realmente tenga A3.
  * En el driver, escala 100% y sin “ajustar al papel”.
- La impresora pide confirmar tamaño → Mismatch trabajo/bandeja.
  * Forzá tamaño correcto y (si hace falta) bandeja en la ruta.
- Márgenes → El código pone márgenes 0 y compensa hard margins.
  * Para “a sangre” real, activá Borderless si el driver lo soporta.
- No encuentra la impresora → printerName debe coincidir exactamente con el nombre en Windows.

Cómo funciona por dentro (mini)
- Acepta Basic (Gray/Color) Print Management + Print Job.
- Resuelve ruta por (Calling AE, Called AE).
- Aplica impresora/duplex/papel/bandeja/fit.
- Renderiza a página completa (usa PageBounds y compensa hard margins).
- Si no forzás papel, intenta usar FilmSizeID del FilmBox.
- Devuelve Success al N-ACTION al encolar el job.


