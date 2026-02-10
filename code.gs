// ============ CONFIGURACIÓN BÁSICA ============

// Nombre de la hoja donde están los datos
const NOMBRE_HOJA   = 'Licencias';
// Carpeta en Drive donde se guardan las tarjetas HTML
const NOMBRE_CARPETA = 'Tarjetas_Licencias_Solutech';
// Zona horaria y formato de fecha
const TZ        = 'America/Panama';
const DATE_FMT  = 'dd/MM/yyyy';

// Correo del que se envían las licencias
const FROM_EMAIL      = 'soporte@solutechpanama.com';
const BCC_LICENCIAS   = 'keys@solutechpanama.com';

// ============ MENÚ ============

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tarjetas Solutech')
    .addItem('Generar tarjeta (fila actual)', 'generarFilaActual')
    .addItem('Generar tarjetas pendientes', 'generarTodasPendientes')
    .addSeparator()
    .addItem('Enviar correo (fila actual)', 'enviarCorreoFilaActual')
    .addSeparator()
    .addItem('Panel de envío', 'abrirPanelEnvio')
    .addToUi();
}

// ============ HELPERS GENERALES ============

function getSheet() {
  return SpreadsheetApp.getActive().getSheetByName(NOMBRE_HOJA) || SpreadsheetApp.getActiveSheet();
}

function getFolder() {
  const it = DriveApp.getFoldersByName(NOMBRE_CARPETA);
  return it.hasNext() ? it.next() : DriveApp.createFolder(NOMBRE_CARPETA);
}

function idx(headers, name) {
  const i = headers.indexOf(name);
  if (i < 0) throw new Error('Falta columna: ' + name);
  return i + 1;
}

function fmtDate(date) {
  if (!date) return '';
  if (!(date instanceof Date)) date = new Date(date);
  return Utilities.formatDate(date, TZ, DATE_FMT);
}

function addYears(date, years) {
  if (!date) return null;
  const d = new Date(date.getTime());
  d.setFullYear(d.getFullYear() + Number(years || 0));
  return d;
}

function cleanPhone(n) {
  if (!n) return '';
  return String(n).replace(/\D+/g, '');
}

// ============ MARCAS / PRODUCTOS ============

function getBrandTheme(productoRaw) {
  const p = (productoRaw || '').toString().toLowerCase();

  // valores por defecto
  let brand = {
    name: 'Genérico',
    color: '#e11d48',
    bg: '#020617',
    logo: 'https://cdn.solutechcloud.com/logos/logo-generic-mail.png',
    activation: 'https://keys.solutechpanama.com/activar'
  };

  if (p.indexOf('mcafee') !== -1) {
    brand = {
      name: 'McAfee',
      color: '#b00014',
      bg: '#020617',
      logo: 'https://cdn.solutechcloud.com/logos/logo-mcafee-mail.png',
      activation: 'https://keys.mcafee.com.pa'
    };
  } else if (p.indexOf('kaspersky') !== -1) {
    brand = {
      name: 'Kaspersky',
      color: '#008c7b',
      bg: '#020617',
      logo: 'https://cdn.solutechcloud.com/logos/logo-kaspersky-mail.png',
      activation: 'https://keys.kasperskypanama.com'
    };
  } else if (p.indexOf('eset') !== -1) {
    brand = {
      name: 'ESET',
      color: '#0099a8',
      bg: '#020617',
      logo: 'https://cdn.solutechcloud.com/logos/logo-eset-mail.png',
      activation: 'https://keys.esetpanama.net'
    };
  } else if (p.indexOf('microsoft') !== -1 || p.indexOf('office') !== -1 || p.indexOf('365') !== -1) {
    brand = {
      name: 'Microsoft 365',
      color: '#f97316',
      bg: '#020617',
      logo: 'https://cdn.solutechcloud.com/logos/logo-m365-mail.png',
      activation: 'https://keys.solutechcloud.com'
    };
  }

  return brand;
}

// ============ ON EDIT: LIMPIAR WHATSAPP / ENVIOS AUTOMÁTICOS ============

function onEdit(e) {
  try {
    if (!e) return;
    limpiarWhatsapp_(e);
    manejarPagoOnEdit_(e);
  } catch (err) {
    // Logger.log(err);
  }
}

function limpiarWhatsapp_(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== NOMBRE_HOJA) return;
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2 || col !== 3) return; // C = WHATSAPP

  const original = e.value || '';
  if (!original) return;

  const cleaned = cleanPhone(original);
  const current = sh.getRange(row, col).getValue();
  if (current === cleaned) return;

  sh.getRange(row, col).setValue(cleaned);
}

function manejarPagoOnEdit_(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== NOMBRE_HOJA) return;

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const cPagado = idx(headers, 'PAGADO');
  const cEstado = idx(headers, 'ESTADO');

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2 || col !== cPagado) return;

  const val = (e.value || '').toString().trim().toLowerCase();
  if (val !== 'si' && val !== 'sí') return; // solo cuando pasa a "Sí"

  const estado = (sh.getRange(row, cEstado).getValue() || '').toString();
  if (estado === 'Enviado') {
    return; // ya se envió antes
  }

  generarTarjetaPorFila_(sh, row);
  enviarCorreoFila_(sh, row);
}

// ============ GENERACIÓN DE TARJETAS ============

function generarFilaActual() {
  const sh = getSheet();
  const row = sh.getActiveRange().getRow();
  if (row < 2) throw new Error('Selecciona una fila con datos (fila 2 o superior).');
  generarTarjetaPorFila_(sh, row);
}

function generarTodasPendientes() {
  const sh = getSheet();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const cEstado = idx(headers, 'ESTADO');
  const lastRow = sh.getLastRow();
  for (let r = 2; r <= lastRow; r++) {
    const est = (sh.getRange(r, cEstado).getValue() || '').toString();
    if (est !== 'Enviado' && est !== 'Generado') {
      generarTarjetaPorFila_(sh, r);
    }
  }
}

function generarTarjetaPorFila_(sh, row) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const cNombre   = idx(headers, 'NOMBRE_CLIENTE');
  const cEmail    = idx(headers, 'EMAIL_CLIENTE');
  const cWhats    = idx(headers, 'WHATSAPP');
  const cProd     = idx(headers, 'PRODUCTO');
  const cDisp     = idx(headers, 'DISPOSITIVOS');
  const cDur      = idx(headers, 'DURACION_ANIOS');
  const cMod      = idx(headers, 'MODALIDAD');
  const cFEmision = idx(headers, 'FECHA_EMISION');
  const cFActiv   = idx(headers, 'FECHA_ACTIVACION');
  const cFVenc    = idx(headers, 'FECHA_VENCIMIENTO');
  const cLic      = idx(headers, 'LICENCIA');
  const cActUrl   = idx(headers, 'ACTIVATION_URL');
  const cLogoUrl  = idx(headers, 'BRAND_LOGO_URL');
  const cInstr    = idx(headers, 'INSTRUCCIONES');
  const cQrUrl    = idx(headers, 'QR_URL');
  const cTarHtml  = idx(headers, 'TARJETA_HTML_URL');
  const cTarMail  = idx(headers, 'TARJETA_EMAIL_HTML');
  const cEstado   = idx(headers, 'ESTADO');

  const nombre   = sh.getRange(row, cNombre).getValue();
  const email    = sh.getRange(row, cEmail).getValue();
  const whatsapp = sh.getRange(row, cWhats).getValue();
  const producto = sh.getRange(row, cProd).getValue();
  const disp     = sh.getRange(row, cDisp).getValue();
  const dur      = sh.getRange(row, cDur).getValue();
  const modalidad= (sh.getRange(row, cMod).getValue() || '').toString();
  let fEmision   = sh.getRange(row, cFEmision).getValue();
  let fActiv     = sh.getRange(row, cFActiv).getValue();
  let fVenc      = sh.getRange(row, cFVenc).getValue();
  const licencia = sh.getRange(row, cLic).getValue();

  if (!fEmision) {
    fEmision = new Date();
    sh.getRange(row, cFEmision).setValue(fEmision);
  }

  // decidir fecha base según modalidad
  let fechaBase = fEmision;
  if (modalidad === 'DESDE_ACTIVACION' && fActiv) {
    fechaBase = fActiv;
  } else if (modalidad === 'DESDE_ACTIVACION' && !fActiv) {
    fechaBase = fEmision;
  }

  fVenc = addYears(fechaBase, dur);
  sh.getRange(row, cFVenc).setValue(fVenc);

  const brand = getBrandTheme(producto);
  sh.getRange(row, cActUrl).setValue(brand.activation);
  sh.getRange(row, cLogoUrl).setValue(brand.logo);

  const instrucciones = 'Visita la página de activación y sigue los pasos en pantalla. Si necesitas ayuda, contáctanos por WhatsApp.';
  sh.getRange(row, cInstr).setValue(instrucciones);

  const qrUrl = 'https://chart.googleapis.com/chart?chs=260x260&cht=qr&chl=' + encodeURIComponent(brand.activation);
  sh.getRange(row, cQrUrl).setValue(qrUrl);

  // plantilla HTML completa (tarjeta bonita para navegador / Drive)
  const htmlCard = construirHtmlTarjeta_({
    nombre,
    email,
    whatsapp,
    producto,
    disp,
    dur,
    modalidad,
    fEmision,
    fActiv,
    fVenc,
    licencia,
    brand,
    qrUrl
  });

  const carpeta = getFolder();
  const nombreArchivo = 'licencia_' + (nombre || 'cliente') + '_' + row + '.html';
  let file = carpeta.createFile(nombreArchivo, htmlCard, MimeType.HTML);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const urlHtml = file.getUrl();
  sh.getRange(row, cTarHtml).setValue(urlHtml);

  // HTML especial para el cuerpo del correo (ligero y compatible)
  const htmlEmail = construirHtmlEmail_({
    nombre,
    producto,
    disp,
    dur,
    fEmision,
    fVenc,
    licencia,
    brand,
    tarjetaUrl: urlHtml,
    activacionUrl: brand.activation
  });
  sh.getRange(row, cTarMail).setValue(htmlEmail);

  sh.getRange(row, cEstado).setValue('Generado');
}

function construirHtmlTarjeta_(data) {
  const {
    nombre,
    producto,
    disp,
    dur,
    fEmision,
    fVenc,
    licencia,
    brand,
    qrUrl
  } = data;

  const titulo = (producto || '') + ' — ' + (disp || 1) + ' dispositivo(s)';
  const fechaEmisionStr = fmtDate(fEmision);
  const fechaVencStr    = fmtDate(fVenc);

  const html = `<!DOCTYPE html>
  <html lang="es">
  <head>
  <meta charset="UTF-8">
  <title>Licencia digital</title>
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <style>
  :root{
    --bg:#020617;
    --card:#020617;
    --muted:#9ca3af;
    --accent:${brand.color};
    --danger:#ef4444;
    --radius:18px;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{
    min-height:100vh;
    display:flex;
    align-items:center;
    justify-content:center;
    font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;
    background:radial-gradient(circle at top, #020617, #000);
    color:#e5e7eb;
    padding:16px;
  }
  .card{
    width:100%;
    max-width:880px;
    background:linear-gradient(135deg,rgba(15,23,42,0.98),rgba(15,23,42,0.96));
    border-radius:24px;
    border:1px solid rgba(148,163,184,0.3);
    box-shadow:0 28px 80px rgba(0,0,0,0.65);
    padding:20px 22px;
  }
  @media (min-width:720px){
    .card{padding:26px 30px;}
    .row{display:flex;gap:18px;align-items:center;}
  }
  .logo-box{
    width:80px;height:80px;
    border-radius:20px;
    background:#020617;
    border:1px solid rgba(148,163,184,0.4);
    display:flex;
    align-items:center;
    justify-content:center;
    overflow:hidden;
  }
  .logo-box img{width:90%;height:auto;object-fit:contain;}
  .title{
    flex:1;
  }
  .title h1{
    font-size:20px;
    margin:0 0 4px;
  }
  .title .meta{
    font-size:13px;
    color:var(--muted);
  }
  .badge{
    display:inline-flex;
    align-items:center;
    gap:6px;
    padding:4px 10px;
    border-radius:999px;
    font-size:11px;
    background:rgba(34,197,94,0.12);
    color:#a7f3d0;
    border:1px solid rgba(52,211,153,0.4);
  }
  .section{
    margin-top:16px;
    padding:14px 16px;
    border-radius:18px;
    background:rgba(15,23,42,0.9);
    border:1px solid rgba(31,41,55,0.9);
  }
  .section h2{
    font-size:13px;
    text-transform:uppercase;
    letter-spacing:.12em;
    color:var(--muted);
    margin-bottom:8px;
  }
  .code-box{
    font-family:Consolas,Monaco,'SF Mono',monospace;
    font-size:16px;
    letter-spacing:0.18em;
    padding:10px 14px;
    border-radius:14px;
    background:#020617;
    border:1px solid rgba(148,163,184,0.5);
    white-space:nowrap;
    overflow-x:auto;
  }
  .actions{
    margin-top:14px;
    display:flex;
    flex-wrap:wrap;
    gap:10px;
  }
  button{
    border-radius:999px;
    border:none;
    font-size:13px;
    padding:8px 16px;
    cursor:pointer;
    display:inline-flex;
    align-items:center;
    gap:6px;
  }
  .btn-primary{
    background:var(--accent);
    color:white;
  }
  .btn-ghost{
    background:transparent;
    color:var(--muted);
    border:1px solid rgba(148,163,184,0.5);
  }
  .footer{
    margin-top:16px;
    font-size:11px;
    color:var(--muted);
  }
  .grid{
    margin-top:16px;
    display:grid;
    grid-template-columns:minmax(0,1.8fr) minmax(0,1.2fr);
    gap:16px;
  }
  @media (max-width:720px){
    .grid{
      grid-template-columns: minmax(0,1fr);
    }
  }
  .qr-box{
    border-radius:20px;
    background:#020617;
    border:1px solid rgba(148,163,184,0.4);
    padding:10px;
    display:flex;
    align-items:center;
    justify-content:center;
  }
  .qr-box img{max-width:100%;height:auto;}
  .small-label{font-size:11px;color:var(--muted);margin-bottom:4px;}
  </style>
  </head>
  <body>
  <article class="card" aria-label="Tarjeta virtual de licencia">
    <div class="row">
      <div class="logo-box">
        <img src="${brand.logo}" alt="${brand.name} logo">
      </div>
      <div class="title">
        <h1>${titulo}</h1>
        <div class="meta">
          Cliente: <strong>${nombre || '(sin nombre)'}</strong> ·
          Emitido: ${fechaEmisionStr || '---'} ·
          Vence: ${fechaVencStr || '---'} ·
          Duración: ${dur || 1} año(s)
        </div>
      </div>
      <div>
        <span class="badge">
          Licencia digital
        </span>
      </div>
    </div>

    <div class="section">
      <h2>Clave de activación</h2>
      <div class="code-box" id="licenciaBox">${licencia || '••••-••••-••••-••••'}</div>
      <div class="actions">
        <button class="btn-ghost" onclick="copiarCodigo()">Copiar código</button>
        <button class="btn-primary" onclick="irActivacion()">Activar ahora</button>
      </div>
    </div>

    <div class="grid">
      <div class="section">
        <h2>Instrucciones</h2>
        <p style="font-size:13px;line-height:1.5;margin-bottom:10px;">
          1. Haz clic en <strong>“Activar ahora”</strong> o abre la página de activación en tu navegador.<br>
          2. Inicia sesión con tu cuenta si es necesario.<br>
          3. Cuando te lo pida, copia y pega esta clave de activación.<br>
          4. Finaliza el proceso y verifica que tu protección esté activa.
        </p>
        <p style="font-size:12px;color:var(--muted);">
          Si tienes dudas o algo no funciona, escríbenos por WhatsApp y con gusto te ayudamos.
        </p>
      </div>

      <div class="section">
        <h2>Activación rápida</h2>
        <div class="small-label">Escanea este código QR</div>
        <div class="qr-box">
          <img src="${qrUrl}" alt="QR de activación">
        </div>
      </div>
    </div>

    <div class="footer">
      Soporte: WhatsApp <strong>+507 6888-6778</strong> · ${FROM_EMAIL}
    </div>
  </article>
  <script>
  function copiarCodigo(){
    const el = document.getElementById('licenciaBox');
    const range = document.createRange();
    range.selectNodeContents(el);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    try{
      document.execCommand('copy');
    }catch(e){}
    sel.removeAllRanges();
  }
  function irActivacion(){
    window.location.href = '${brand.activation}';
  }
  </script>
  </body>
  </html>`;
    return html;
}

// ============ ENVÍO DE CORREO HTML ============

function construirHtmlEmail_(data) {
  const {
    nombre,
    producto,
    disp,
    dur,
    fEmision,
    fVenc,
    licencia,
    brand,
    tarjetaUrl,
    activacionUrl
  } = data;

  const fechaEmisionStr = fmtDate(fEmision);
  const fechaVencStr    = fmtDate(fVenc);
  const tituloProducto  = (producto || 'Tu licencia digital') +
                          (disp ? ' — ' + disp + ' dispositivo(s)' : '');

  // Bloque de logo (usa brand.logo y brand.name)
  const logoHtml = (brand && brand.logo)
    ? '      <div style="text-align:center;margin:0 0 16px;">'
      + '        <img src="' + brand.logo + '" '
      + 'alt="' + (brand.name || 'Logo') + '" '
      + 'style="max-width:120px;height:auto;display:inline-block;'
      + 'border-radius:12px;">'
      + '      </div>'
    : '';


  return ''
    + '<div style="background-color:#020617;padding:24px 0;'
    + 'font-family:system-ui,Arial,sans-serif;">'
    + '  <div style="max-width:640px;margin:0 auto;color:#e5e7eb;">'

    + '    <p style="font-size:14px;margin:0 0 16px;">'
    + 'Hola ' + (nombre || '') + ',</p>'

    + '    <p style="font-size:14px;margin:0 0 16px;">'
    + 'Gracias por tu compra. Aquí tienes los datos de tu licencia digital:'
    + '    </p>'

    + '    <div style="background:#020617;border-radius:16px;'
    + '                border:1px solid #1f2937;padding:20px;">'

    // Logo de la marca
    + logoHtml

    + '      <p style="font-size:13px;color:#9ca3af;margin:0 0 4px;">Producto</p>'
    + '      <p style="font-size:15px;font-weight:600;margin:0 0 12px;">'
    +          tituloProducto + '</p>'

    + '      <p style="font-size:12px;color:#9ca3af;margin:0 0 4px;">'
    + 'Código de licencia</p>'
    + '      <p style="font-size:18px;font-weight:600;'
    + '                font-family:Consolas,monospace;'
    + '                margin:0 0 10px;">'
    +          (licencia || '') + '</p>'

    + '      <p style="font-size:12px;color:#9ca3af;margin:0 0 4px;">'
    + 'Emitido: ' + (fechaEmisionStr || '---')
    + ' · Vence: ' + (fechaVencStr || '---')
    + ' · Duración: ' + (dur || 1) + ' año(s)'
    + '      </p>'

    + (activacionUrl
        ? '      <p style="margin:14px 0 0;">'
          + '        <a href="' + activacionUrl + '" '
          + 'style="display:inline-block;background:#ef4444;'
          + 'color:#ffffff;text-decoration:none;font-size:14px;'
          + 'padding:10px 18px;border-radius:999px;">'
          + 'Activar ahora</a></p>'
        : '')

    + (tarjetaUrl
        ? '      <p style="font-size:12px;color:#9ca3af;margin:12px 0 0;">'
          + 'También puedes abrir la tarjeta completa en este enlace: '
          + '<a href="' + tarjetaUrl + '" style="color:#93c5fd;">ver tarjeta</a>'
          + '</p>'
        : '')

    + '    </div>'

    + '    <p style="font-size:12px;color:#9ca3af;margin:16px 0 0;">'
    + 'Soporte: WhatsApp +507 6888-6778 · '
    + '<a href="mailto:soporte@solutechpanama.com" '
    + 'style="color:#93c5fd;">soporte@solutechpanama.com</a>'
    + '    </p>'

    + '  </div>'
    + '</div>';
}


// ============ ENVÍO DE CORREO ============

function enviarCorreoFilaActual() {
  const sh = getSheet();
  const row = sh.getActiveRange().getRow();
  if (row < 2) throw new Error('Selecciona una fila con datos.');
  enviarCorreoFila_(sh, row);
}

function enviarCorreoFila_(sh, row) {
  const headers   = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const cNombre   = idx(headers, 'NOMBRE_CLIENTE');
  const cEmail    = idx(headers, 'EMAIL_CLIENTE');
  const cProd     = idx(headers, 'PRODUCTO');
  const cLic      = idx(headers, 'LICENCIA');
  const cTarMail  = idx(headers, 'TARJETA_EMAIL_HTML');
  const cTarHtml  = idx(headers, 'TARJETA_HTML_URL');
  const cEstado   = idx(headers, 'ESTADO');
  const cEnviado  = idx(headers, 'ENVIADO');
  const cEnvFecha = idx(headers, 'ENVIADO_FECHA');

  const nombre  = sh.getRange(row, cNombre).getValue();
  const email   = sh.getRange(row, cEmail).getValue();
  const prod    = sh.getRange(row, cProd).getValue();
  const licencia= sh.getRange(row, cLic).getValue();
  const htmlTar = sh.getRange(row, cTarMail).getValue();
  const urlTar  = sh.getRange(row, cTarHtml).getValue();

  if (!email) throw new Error('La fila no tiene EMAIL_CLIENTE.');
  if (!licencia) throw new Error('La fila no tiene LICENCIA.');
  if (!htmlTar) throw new Error('La tarjeta no ha sido generada aún.');

  const asunto = 'Tu licencia digital — ' + (prod || 'Producto Solutech');

  const cuerpoTexto = 'Hola ' + (nombre || '') + ',\n\n' +
    'Gracias por tu compra. Adjuntamos tu tarjeta de licencia digital.\n\n' +
    'Código de licencia: ' + licencia + '\n' +
    'También puedes abrir la tarjeta en línea en este enlace: ' + urlTar + '\n\n' +
    'Cualquier duda, escríbenos por WhatsApp +507 6888-6778.\n\n' +
    'Solutech Panamá';

  // Usamos directamente el HTML ya preparado para email (ligero)
  const cuerpoHtml = htmlTar;

  MailApp.sendEmail({
    to: email,
    name: 'Solutech Panamá',
    from: FROM_EMAIL,
    bcc: BCC_LICENCIAS,
    subject: asunto,
    body: cuerpoTexto,
    htmlBody: cuerpoHtml
  });

  sh.getRange(row, cEstado).setValue('Enviado');
  sh.getRange(row, cEnviado).setValue('Sí');
  sh.getRange(row, cEnvFecha).setValue(new Date());
}

// ============ PANEL HTML (DATOS DE FILA) ============

function abrirPanelEnvio() {
  const sh  = getSheet();
  const row = sh.getActiveRange().getRow();
  let datos = null;

  if (row >= 2) {
    const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const cNombre   = idx(headers, 'NOMBRE_CLIENTE');
    const cEmail    = idx(headers, 'EMAIL_CLIENTE');
    const cProd     = idx(headers, 'PRODUCTO');
    const cLic      = idx(headers, 'LICENCIA');
    const cEstado   = idx(headers, 'ESTADO');
    const cPagado   = idx(headers, 'PAGADO');
    const cEnviado  = idx(headers, 'ENVIADO');

    datos = {
      row,
      nombre:   sh.getRange(row, cNombre).getValue(),
      email:    sh.getRange(row, cEmail).getValue(),
      producto: sh.getRange(row, cProd).getValue(),
      licencia: sh.getRange(row, cLic).getValue(),
      estado:   sh.getRange(row, cEstado).getValue(),
      pagado:   sh.getRange(row, cPagado).getValue(),
      enviado:  sh.getRange(row, cEnviado).getValue()
    };
  }

  const tpl = HtmlService.createTemplateFromFile('panelEnvio');
  tpl.datos = datos;

  const html = tpl.evaluate()
    .setTitle('Envío de licencias')
    .setWidth(320);

  SpreadsheetApp.getUi().showSidebar(html);
}

function enviarDesdePanel() {
  const sh = getSheet();
  const row = sh.getActiveRange().getRow();
  if (row < 2) throw new Error('Selecciona una fila con datos.');
  generarTarjetaPorFila_(sh, row); // por si aún no se ha generado
  enviarCorreoFila_(sh, row);
  return 'Licencia enviada al cliente.';
}


// ===================== FORMULARIO PÚBLICO CLIENTE (CREA FILA NUEVA) =====================

function doGet() {
  return HtmlService.createTemplateFromFile('formClienteNuevo')
    .evaluate()
    .setTitle('Solicitud de licencia Solutech')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Crea una fila nueva en la hoja "Licencias".
 * El cliente solo llena datos básicos; tú completas luego licencia/pago/etc.
 */
function registrarSolicitud(payload) {
  const sh = getSheet();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  const cNombre = idx(headers, 'NOMBRE_CLIENTE');
  const cEmail  = idx(headers, 'EMAIL_CLIENTE');
  const cWhats  = idx(headers, 'WHATSAPP');
  const cObs    = idx(headers, 'OBSERVACIONES');   // Col U
  const cProd   = idx(headers, 'PRODUCTO');
  const cDisp   = idx(headers, 'DISPOSITIVOS');
  const cDur    = idx(headers, 'DURACION_ANIOS');

  const cEstado = idx(headers, 'ESTADO');
  const cPagado = idx(headers, 'PAGADO');

  const row = sh.getLastRow() + 1;

  // Normaliza
  const nombre = (payload.nombre || '').toString().trim();
  const email  = (payload.email  || '').toString().trim();
  const whatsapp = cleanPhone(payload.whatsapp || '');
  const observaciones = (payload.observaciones || '').toString().trim();

  const producto = (payload.producto || '').toString().trim();
  const dispositivos = (payload.dispositivos || '').toString().trim();
  const anios = (payload.anios || '').toString().trim();

  // Requeridos mínimos
  if (!nombre) throw new Error('Falta el nombre.');
  if (!email)  throw new Error('Falta el email.');

  // Escribir campos
  sh.getRange(row, cNombre).setValue(nombre);
  sh.getRange(row, cEmail ).setValue(email);
  if (whatsapp) sh.getRange(row, cWhats).setValue(whatsapp);
  if (observaciones) sh.getRange(row, cObs).setValue(observaciones);

  if (producto) sh.getRange(row, cProd).setValue(producto);
  if (dispositivos) sh.getRange(row, cDisp).setValue(dispositivos);
  if (anios) sh.getRange(row, cDur).setValue(anios);

  // Defaults
  sh.getRange(row, cEstado).setValue('Pendiente');
  sh.getRange(row, cPagado).setValue('No');

  return { ok: true, row };
}


function validarTurnstile_(token) {
  const secret = PropertiesService.getScriptProperties().getProperty('0x4AAAAAACaNg7x5TlbZ2I7hJ84nhNFWsi8');
  if (!secret) throw new Error('Falta TURNSTILE_SECRET en Script Properties.');
  if (!token) return false;

  const url = 'https://challenges.cloudflare.com/turnstile/v0/siteverify';

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    payload: {
      secret: secret,
      response: token
    },
    muteHttpExceptions: true
  });

  const data = JSON.parse(res.getContentText() || '{}');
  return !!data.success;
}

function doGet() {
  return HtmlService.createTemplateFromFile('formClienteNuevo')
    .evaluate()
    .setTitle('Solicitud de licencia Solutech')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function registrarSolicitud(datos) {
  // 1) Validar Turnstile (obligatorio)
  const ok = validarTurnstile_(datos && datos.turnstile);
  if (!ok) throw new Error('No se pudo validar el verificador. Intenta nuevamente.');

  // 2) Validaciones mínimas
  const nombre = (datos.nombre || '').toString().trim();
  const email  = (datos.email  || '').toString().trim();
  const whatsapp = cleanPhone(datos.whatsapp || '');
  const observaciones = (datos.observaciones || '').toString().trim();

  const producto = (datos.producto || '').toString().trim();
  const dispositivos = (datos.dispositivos || '').toString().trim();
  const anios = (datos.anios || '').toString().trim();

  if (!nombre) throw new Error('Falta el nombre.');
  if (!email) throw new Error('Falta el email.');
  if (!whatsapp) throw new Error('Falta el WhatsApp.');

  // 3) Insertar nueva fila
  const sh = getSheet();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  const cNombre = idx(headers, 'NOMBRE_CLIENTE');   // A
  const cEmail  = idx(headers, 'EMAIL_CLIENTE');    // B
  const cWhats  = idx(headers, 'WHATSAPP');         // C
  const cProd   = idx(headers, 'PRODUCTO');         // D
  const cDisp   = idx(headers, 'DISPOSITIVOS');     // E
  const cDur    = idx(headers, 'DURACION_ANIOS');   // F
  const cEstado = idx(headers, 'ESTADO');           // R
  const cPagado = idx(headers, 'PAGADO');           // S
  const cObs    = idx(headers, 'OBSERVACIONES');    // U

  const row = sh.getLastRow() + 1;

  sh.getRange(row, cNombre).setValue(nombre);
  sh.getRange(row, cEmail).setValue(email);
  sh.getRange(row, cWhats).setValue(whatsapp);

  if (producto) sh.getRange(row, cProd).setValue(producto);
  if (dispositivos) sh.getRange(row, cDisp).setValue(dispositivos);
  if (anios) sh.getRange(row, cDur).setValue(anios);

  if (observaciones) sh.getRange(row, cObs).setValue(observaciones);

  sh.getRange(row, cEstado).setValue('Pendiente');
  sh.getRange(row, cPagado).setValue('No');

  return { ok: true, row };
}


// ============ FORMULARIO PÚBLICO (WEB APP) + TURNSTILE ============
// 1) Crea en Script Properties:
//    TURNSTILE_SECRET = (tu secret key de Cloudflare Turnstile)
//    (opcional) ALLOW_ORIGIN = https://tu-dominio.com  (o *)
// 2) Deploy: Implementar > Nueva implementación > Aplicación web
//    Ejecutar como: Tú
//    Quién tiene acceso: Cualquiera (o Cualquiera con el enlace)

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) ? e.postData.contents : '{}');

    // Validación básica
    const nombre = (payload.nombre || '').toString().trim();
    const email  = (payload.email || '').toString().trim();
    const whatsapp = (payload.whatsapp || '').toString().trim();
    if (!nombre || !email || !whatsapp) {
      return json_(false, 'Faltan campos obligatorios (nombre, email, whatsapp).');
    }

    // Turnstile
    const token = (payload.turnstileToken || '').toString().trim();
    const okCaptcha = verificarTurnstile_(token);
    if (!okCaptcha) {
      return json_(false, 'Verificación fallida. Intenta de nuevo.');
    }

    // Registrar en Sheet
    const res = registrarSolicitudPublica_(payload);
    return json_(true, 'OK', res);

  } catch (err) {
    return json_(false, 'Error: ' + (err && err.message ? err.message : err));
  }
}

function registrarSolicitudPublica_(p) {
  const sh = getSheet();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  // Columnas existentes en tu Sheet (según tu sistema actual)
  const cNombre = idx(headers, 'NOMBRE_CLIENTE');
  const cEmail  = idx(headers, 'EMAIL_CLIENTE');
  const cWhats  = idx(headers, 'WHATSAPP');
  const cProd   = idx(headers, 'PRODUCTO');
  const cDisp   = idx(headers, 'DISPOSITIVOS');
  const cAnios  = idx(headers, 'DURACION_ANIOS');
  const cObs    = idx(headers, 'OBSERVACIONES'); // tu columna U

  // Opcionales para setear “bonito” desde el inicio (si existen)
  const cEstado = headers.includes('ESTADO') ? idx(headers, 'ESTADO') : null;
  const cPagado = headers.includes('PAGADO') ? idx(headers, 'PAGADO') : null;
  const cEnviado= headers.includes('ENVIADO') ? idx(headers, 'ENVIADO') : null;

  const nextRow = sh.getLastRow() + 1;

  // Normalizaciones
  const nombre = (p.nombre || '').toString().trim();
  const email  = (p.email || '').toString().trim();
  const whatsapp = cleanPhone(p.whatsapp || '');
  const producto  = (p.producto || '').toString().trim();
  const dispositivos = (p.dispositivos || '').toString().trim();
  const anios = (p.anios || '').toString().trim();
  const obs  = (p.observaciones || '').toString().trim();

  sh.getRange(nextRow, cNombre).setValue(nombre);
  sh.getRange(nextRow, cEmail).setValue(email);
  sh.getRange(nextRow, cWhats).setValue(whatsapp);
  sh.getRange(nextRow, cProd).setValue(producto);
  if (dispositivos) sh.getRange(nextRow, cDisp).setValue(Number(dispositivos));
  if (anios)        sh.getRange(nextRow, cAnios).setValue(Number(anios));
  if (obs)          sh.getRange(nextRow, cObs).setValue(obs);

  // Estados iniciales (si existen)
  if (cEstado) sh.getRange(nextRow, cEstado).setValue('Pendiente');
  if (cPagado) sh.getRange(nextRow, cPagado).setValue('No');
  if (cEnviado) sh.getRange(nextRow, cEnviado).setValue('No');

  return { row: nextRow };
}

function verificarTurnstile_(token) {
  if (!token) return false;

  const secret = PropertiesService.getScriptProperties().getProperty('0x4AAAAAACaNg7x5TlbZ2I7hJ84nhNFWsi8');
  if (!secret) throw new Error('Falta TURNSTILE_SECRET en Script Properties.');

  const resp = UrlFetchApp.fetch('https://challenges.cloudflare.com/turnstile/v0/siteverify', {
    method: 'post',
    payload: {
      secret: secret,
      response: token
    },
    muteHttpExceptions: true
  });

  const data = JSON.parse(resp.getContentText() || '{}');
  return !!data.success;
}

function json_(ok, message, extra) {
  const allow = PropertiesService.getScriptProperties().getProperty('ALLOW_ORIGIN') || '*';

  const out = ContentService
    .createTextOutput(JSON.stringify({ ok, message, ...(extra ? extra : {}) }))
    .setMimeType(ContentService.MimeType.JSON);

  // Headers CORS (Apps Script sí permite setHeader en TextOutput)
  out.setHeader('Access-Control-Allow-Origin', allow);
  out.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  out.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  return out;
}
