// @ts-nocheck
const hojasPermitidas = ["EnvíoAsignación", "EnvíoAsignación2", "EnvíoAsignación3"];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Asignación")
    .addItem("📧 Enviar Correo", "enviarCorreoAsignacion")
    .addItem("🚮 Limpiar", "limpiarAsignacion")
    .addToUi();
}

function onEdit(e) {
 const ss = e.source;
 const hoja = ss.getActiveSheet();
 if (!hojasPermitidas.includes(hoja.getName())) return;

  // 🔹 1) Limpiar saltos de línea en lo pegado
  let valor = e.range.getValue();
  if (typeof valor === "string") {
    valor = valor.replace(/\r?\n|\r/g, " ");
    e.range.setValue(valor);
  }

  const lastRow = hoja.getLastRow();

  // 🔹 2) Aplicar formato general (EXCLUYE encabezado para no perder estilos)
  if (lastRow > 1) {
    const rango = hoja.getRange("A2:AX" + lastRow);
    rango.setFontFamily("Calibri")
         .setFontSize(12)
         .setHorizontalAlignment("center")
         .setVerticalAlignment("middle")
         .setWrap(true)
         .setFontColor("black")
         .setFontWeight("normal")
         .setBackground(null);
  }

  // 🔹 3) Recalcular SI/NO/BLANCO de fotos en AX
  const hojaLista = ss.getSheetByName("Lista");
  if (!hojaLista) return;

  const clientesLista = hojaLista.getRange("A2:A" + hojaLista.getLastRow())
    .getValues().flat().filter(x => x);

  const clientes = hoja.getRange(2, 3, lastRow - 1).getValues().flat(); // col C
  const resultados = clientes.map(c => {
    if (!c) return [""]; // vacío → nada en AX
    if (clientesLista.includes(c)) return ["SI"];
    return ["NO"];
  });
  if (resultados.length > 0) {
    hoja.getRange(2, 50, resultados.length).setValues(resultados); // col AX
  }

  // 🔹 4) Formatos especiales
  if (lastRow > 1) {
    const rangoAA = hoja.getRange(2, 27, lastRow - 1); // Cita Retiro
    const rangoAF = hoja.getRange(2, 32, lastRow - 1); // Posicionamiento
    const rangoAX = hoja.getRange(2, 50, lastRow - 1); // Fotos
    const rangoQ = hoja.getRange(2, 17, lastRow - 1); // Gasificado

    // AA → fondo azul
    rangoAA.setBackground("#9fc5e8").setFontWeight("bold");
    
    // AF → fondo amarillo
    rangoAF.setBackground("#ffe599").setFontWeight("bold");
    
    rangoQ.setBackground ( "#FF0000").setFontWeight("bold");

    // AX → SI/NO colores
    const valoresAX = rangoAX.getValues();
    valoresAX.forEach((fila, i) => {
      const celda = rangoAX.getCell(i + 1, 1);
      if (fila[0] === "SI") {
        celda.setBackground("#c6efce")
             .setFontColor("#0d6a0e")
             .setFontWeight("bold");
      } else if (fila[0] === "NO") {
        celda.setBackground("#ffc7ce")
             .setFontColor("#a20c13")
             .setFontWeight("bold");
      } else {
        celda.setBackground(null).setFontColor("black").setFontWeight("normal"); // vacío
      }
    });
  }

  // 🔹 5) Encabezado fijo (A1:AX1)
  const encabezado = hoja.getRange("A1:AX1");
  encabezado.setBackground("#00499b")
            .setFontColor("white")
            .setFontWeight("bold")
            .setFontFamily("Calibri")
            .setFontSize(12)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");

  // 🔹 6) Dibujar bordes en TODAS las filas según C o G
  if (lastRow > 1) {
    const valoresC = hoja.getRange(2, 3, lastRow - 1).getValues().flat(); // col C
    const valoresG = hoja.getRange(2, 7, lastRow - 1).getValues().flat(); // col G

    for (let i = 0; i < lastRow - 1; i++) {
      const fila = i + 2; // empieza en fila 2
      const rangoFila = hoja.getRange(fila, 1, 1, hoja.getLastColumn());

      if (valoresC[i] || valoresG[i]) {
        rangoFila.setBorder(true, true, true, true, true, true, "#555555", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      } else {
        rangoFila.setBorder(false, false, false, false, false, false);
      }
    }
  }
}

function enviarCorreoAsignacion() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const hoja = ss.getActiveSheet();
 if (!hojasPermitidas.includes(hoja.getName())) {
  SpreadsheetApp.getUi().alert("Esta herramienta solo funciona en las hojas permitidas.");
  return;
 }

  const hojaDestinatarios = ss.getSheetByName("Destinatarios");
  const hojaConfig = ss.getSheetByName("Configuración");

  const transportistaRaw = hoja.getRange("V2").getValue();
  const valorFecha = hoja.getRange("AF2").getValue();

  if (!transportistaRaw || !valorFecha) {
    SpreadsheetApp.getUi().alert("Faltan datos: Transportista o Fecha de posicionamiento.");
    return;
  }

  let fecha = valorFecha instanceof Date ? valorFecha : null;
  if (!fecha) {
    SpreadsheetApp.getUi().alert("La fecha de cita de retiro no es válida.");
    return;
  }

  // ✅ Quitar RUC al transportista (cualquier número al final)
  const transportista = String(transportistaRaw).replace(/\s+\d+$/, "");

  // ✅ Validar puerto (columna F)
  const lastRow = hoja.getLastRow();
  const puertos = hoja.getRange(2, 6, lastRow - 1).getValues().flat().filter(x => x);
  const puertoUnico = [...new Set(puertos)];

  if (puertoUnico.length !== 1) {
    SpreadsheetApp.getUi().alert("Los servicios tienen diferentes puertos en columna F. Deben ser todos iguales.");
    return;
  }
  const puerto = puertoUnico[0].toString().toUpperCase(); // Ej: CALLAO, CHANCAY, PISCO

  // Buscar correos de destinatarios
  const listaDest = hojaDestinatarios.getDataRange().getValues();
  const destinatarios = listaDest.find(row => row[0] === transportistaRaw)?.[1];
  if (!destinatarios) {
    SpreadsheetApp.getUi().alert("No se encontraron correos para el transportista.");
    return;
  }

  // Fecha en texto
  const dias = ["Domingo","Lunes","Martes","Miércoles","Jueves","Viernes","Sábado"];
  const meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
  const fechaTexto = ${dias[fecha.getDay()]} ${fecha.getDate()} de ${meses[fecha.getMonth()]};

  // Validar remitente
  const correoUsuario = Session.getActiveUser().getEmail();
  const dataConfig = hojaConfig.getRange("A2:E").getValues();
  const filaRemitente = dataConfig.find(row => row[0] === correoUsuario);
  if (!filaRemitente) {
    SpreadsheetApp.getUi().alert("No estás autorizado a enviar este correo.");
    return;
  }
  const [_, nombre, puesto, telefono, correo] = filaRemitente;

  // 🔹 Columnas que se enviarán (con AA colocado después de H Contenedor)
  const columnas = [
    {letra:"C", idx:2},   // Cliente
    {letra:"E", idx:4},   // Producto
    {letra:"F", idx:5},   // Puerto Embarque
    {letra:"G", idx:6},   // Booking
    {letra:"H", idx:7},   // Contenedor
    {letra:"AA", idx:26}, // Cita Retiro (se muestra aquí)
    {letra:"I", idx:8},   // Nave
    {letra:"O", idx:14},  // ColdTreatment
    {letra:"Q", idx:16},  // Gasificado
    {letra:"S", idx:18},  // Pre enfriado
    {letra:"V", idx:21},  // Transp. Dueño
    {letra:"W", idx:22},  // Tracto
    {letra:"X", idx:23},  // Carreta
    {letra:"Y", idx:24},  // Chofer
    {letra:"Z", idx:25},  // Deposito Retiro
    {letra:"AF", idx:31}, // Posicionamiento
    {letra:"AH", idx:33}, // Planta
    {letra:"AO", idx:40}, // Deposito Retorno
    {letra:"AT", idx:45}, // Zona
    {letra:"AX", idx:49}  // Fotos (SI/NO)
  ];

  // Encabezado
  const encabezado = columnas.map(c => hoja.getRange(1, c.idx+1).getValue());

  // Datos
  const datos = hoja.getRange(2,1,lastRow-1,50).getValues()
    .filter(row => row.some(celda => celda !== "" && celda !== null));

  // Construir tabla HTML
  let html = <table border='1' cellpadding='5' style='border-collapse:collapse; font-family:Calibri; font-size:12px; text-align:center;'>;
  html += <tr style='background:#00499b; color:white; font-weight:bold;'>;
  encabezado.forEach(t => html += <th>${t}</th>);
  html += </tr>;

  datos.forEach(row => {
    html += <tr>;
    columnas.forEach((col) => {
      let valor = row[col.idx];
      let estilo = "color:#003366; font-size:12px; font-family:Calibri;";

      if (valor instanceof Date) valor = formatearFecha(valor);

      if (col.letra === "AA") {
        estilo += "font-weight:bold;background:#9fc5e8;";
      } else if (col.letra === "AF") {
        estilo += "font-weight:bold;background:#ffe599;";
        } else if (col.letra === "Q"){
		    estilo += "font-weight:bold;background:#FF0000;";
      } else if (col.letra === "AX") {
        if (valor === "SI") {
          estilo += "font-weight:bold;background:#c6efce;color:#0d6a0e;";
        } else if (valor === "NO") {
          estilo += "font-weight:bold;background:#ffc7ce;color:#a20c13;";
        }
      }

      html += <td style='${estilo}'>${valor ?? ""}</td>;
    });
    html += </tr>;
  });
  html += </table>;

  // Cuerpo del correo con logo antes de la firma (igual al tuyo)
  const cuerpo = `
  <div style="font-family:Calibri; font-size:14px;">
<p>Estimados,</p>
    <p>Les adjunto la programación detallada para el día <b>${fechaTexto}</b>.</p>
    <p>Por favor, tomen en cuenta la hora de cita de retiro, la cual está resaltada en amarillo. Es importante que envíen los datos de la unidad asignada para cada servicio.</p>
    <p>Si necesitan realizar algún cambio operativo (como cambio de tracto, carreta o conductor), les pedimos que lo informen por este medio y también en el grupo de WhatsApp, haciendo referencia a la reserva o Booking correspondiente.</p>
    <p>Adicionalmente, les pido que consideren las siguientes observaciones al momento de realizar el servicio:</p>
    <ul>
      <li><b>Horas de cita sugerida:</b> Brindar las horas de citas full lo más pronto posible, a fin de gestionar citas con anticipación y no incurrir en sobrecostos. New Transport no se hace responsable de algún sobrecosto de no tener la información a su debido momento.</li>
      <li><b>Informar retrasos:</b> En caso de demoras o retrasos, es fundamental que informen la hora estimada de llegada a la planta.</li>
      <li><span style="background-color: #a9f9ff;"><b>Servicios congelados:</b> Los servicios de productos congelados deben llegar pre-enfriados. Para ello, envíen fotos del encendido del equipo 2 horas antes del posicionamiento y también a su llegada a la planta.</span></li>
      <li><b>Gasificado/Sini/Controlador:</b> Se les informará si aplica y siempre revisen el control de embarque donde se indicará con un sello.</li>
      <li><b>Bolsas de seguridad:</b> Es crucial que las bolsas de seguridad no sean abiertas hasta su entrega en planta.</li>
      <li><b>Fotos de precintado y guías:</b> Por favor, envíen fotos del precintado, GRR, GRT antes de salir de planta y esperar nuestra confirmación para iniciar ruta a puerto. Recordar que la unidad no puede salir de planta sin el canal correspondiente.</li>
      <li><span style="background-color: #ffffa7;"><b>Reporte de GPS:</b> Es indispensable enviar un reporte de GPS, monitoreo e incidencias cada 2 horas desde el inicio hasta la finalización del servicio. La falta de reportes con horarios detallados implicará que no se asumirá sobrestadía en caso de demoras en planta, retiros o entregas full.</span></li>
    </ul>
    <p>El cumplimiento de estas observaciones nos ayudará a evitar sobrecostos y afectaciones en el servicio.</p>
    <br>${html}<br><br>
    <div>
      <img src="http://www.newtransport.net/img/logo.png" height="80"><br><br>
      <b>${nombre}</b><br>
      ${puesto}<br>
      New Transport S.A.<br>
      Teléfono: ${telefono}<br>
      ${correo}<br>
      www.newtransport.net
    </div>
  </div>`;

  // ✅ Asunto con puerto dinámico y transportista sin RUC, todo en mayúsculas
  const asunto = EMBARQUES POR ${puerto} || PROGRAMACION ${fechaTexto} - ${transportista} || FLOTA NT.toUpperCase();

  GmailApp.sendEmail(destinatarios, asunto, "", { htmlBody: cuerpo });
  SpreadsheetApp.getUi().alert("Correo enviado correctamente.");
}

function formatearFecha(fecha) {
  if (!fecha || !(fecha instanceof Date)) return fecha || "";
  return Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
}

function limpiarAsignacion() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const hoja = ss.getActiveSheet();

 if (!hojasPermitidas.includes(hoja.getName())) {
  SpreadsheetApp.getUi().alert("Esta opción solo funciona en las hojas permitidas.");
  return;
 }

  const lastRow = hoja.getLastRow();
  if (lastRow <= 1) return;

  hoja.getRange(2, 1, lastRow - 1, hoja.getLastColumn()).clearContent().clearFormat();

  //SpreadsheetApp.getUi().alert("Se limpiaron los datos correctamente.");
