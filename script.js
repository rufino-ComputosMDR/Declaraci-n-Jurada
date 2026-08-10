let datosExcelRaw = [];
let datosExcel = [];
let nombresColumnas = [];
let filaActual = 0;

function col2idx(colStr) {
  let str = colStr.toUpperCase();
  let sum = 0;
  for (let i = 0; i < str.length; i++) {
    sum *= 26;
    sum += str.charCodeAt(i) - 64;
  }
  return sum - 1;
}

function obtenerIndiceColumna(palabrasClave) {
  for (let i = 0; i < nombresColumnas.length; i++) {
    const colNombre = (nombresColumnas[i] || "").toLowerCase().trim();
    for (let palabra of palabrasClave) {
      if (colNombre.includes(palabra.toLowerCase())) {
        return i;
      }
    }
  }
  return -1;
}

function getValByKeywords(fila, palabrasClave) {
  const idx = obtenerIndiceColumna(palabrasClave);
  return (idx !== -1 && fila[idx] !== undefined) ? fila[idx] : "";
}

window.addEventListener('DOMContentLoaded', () => {
  fetch('declara.xlsx')
    .then(response => {
      if (!response.ok) throw new Error("No se encuentra declara.xlsx");
      return response.arrayBuffer();
    })
    .then(data => {
      const workbook = XLSX.read(data, { type: 'array' });
      const primeraHoja = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[primeraHoja];
      
      datosExcelRaw = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

      if (datosExcelRaw.length > 1) {
        let filaEncabezadoIdx = 0;
        for (let i = 0; i < datosExcelRaw.length; i++) {
          const tieneTexto = datosExcelRaw[i].some(celda => String(celda).trim() !== "");
          if (tieneTexto) {
            filaEncabezadoIdx = i;
            break;
          }
        }

        const cabecerasOriginales = datosExcelRaw[filaEncabezadoIdx];

        for (let i = 0; i <= col2idx('BI'); i++) {
          nombresColumnas[i] = (cabecerasOriginales[i] && cabecerasOriginales[i].toString().trim() !== "") 
            ? cabecerasOriginales[i].toString().trim() 
            : `Columna ${i + 1}`;
        }

        const filasProcesadas = datosExcelRaw.slice(filaEncabezadoIdx + 1).filter(row => {
          return row.some(celda => celda !== undefined && celda !== null && String(celda).trim() !== "");
        });

        datosExcel = filasProcesadas.map(row => {
          let filaObj = {};
          for (let i = 0; i <= col2idx('BI'); i++) {
            filaObj[i] = (row[i] !== undefined) ? row[i] : "";
          }
          return filaObj;
        });

        datosExcel.sort((a, b) => {
          const rawA = getValByKeywords(a, ['agente legajo', 'legajo']);
          const rawB = getValByKeywords(b, ['agente legajo', 'legajo']);
          
          const legA = parseInt(rawA);
          const legB = parseInt(rawB);

          const esNumA = !isNaN(legA);
          const esNumB = !isNaN(legB);

          if (esNumA && esNumB) return legA - legB;
          if (esNumA && !esNumB) return -1;
          if (!esNumA && esNumB) return 1;
          return 0;
        });

        poblarSelector();
        construirTablaModal();
        mostrarFila(0);
      } else {
        alert("El archivo declara.xlsx está vacío.");
      }
    })
    .catch(error => {
      console.error("Error al cargar declara.xlsx:", error);
      alert("Error: Verifica que 'declara.xlsx' esté en la misma carpeta.");
    });
});

/* POBLAR SELECTOR INCLUYENDO EL AÑO DE PRESENTACIÓN */
function poblarSelector() {
  const select = document.getElementById('selectorFilas');
  select.innerHTML = '';

  datosExcel.forEach((fila, index) => {
    const option = document.createElement('option');
    option.value = index;
    
    const legajo = getValByKeywords(fila, ['agente legajo', 'legajo']) || 'S/L';
    const nombre = getValByKeywords(fila, ['nombre y apellido', 'agente', 'nombre']) || 'Sin Nombre';
    const dni = getValByKeywords(fila, ['dni', 'documento']) || 'S/D';
    
    // Extracción de fecha y año
    const fechaRaw = getValByKeywords(fila, ['registrado', 'marca temporal', 'fecha']);
    const fechaFormateada = formatearFecha(fechaRaw);
    
    let anio = 'S/A';
    if (fechaFormateada && fechaFormateada !== '-') {
      const partes = fechaFormateada.split('/');
      if (partes.length === 3) {
        anio = partes[2].substring(0, 4);
      }
    }

    option.textContent = `[Año: ${anio}] - Legajo: ${legajo} - ${nombre} (DNI: ${dni})`;
    select.appendChild(option);
  });
}

function mostrarFila(index) {
  if (index < 0 || index >= datosExcel.length) return;
  
  filaActual = parseInt(index);
  const fila = datosExcel[filaActual];

  document.getElementById('selectorFilas').value = filaActual;
  document.getElementById('contador').textContent = `Fila ${filaActual + 1} de ${datosExcel.length}`;
  document.getElementById('btnPrev').disabled = (filaActual === 0);
  document.getElementById('btnNext').disabled = (filaActual === datosExcel.length - 1);

  const fechaRealizacion = formatearFechaHora(getValByKeywords(fila, ['registrado', 'marca temporal', 'fecha']));
  const legajo = getValByKeywords(fila, ['agente legajo', 'legajo']);
  const nombre = getValByKeywords(fila, ['nombre y apellido', 'agente']);
  const domicilio = getValByKeywords(fila, ['domicilio actual', 'domicilio']);
  const dni = getValByKeywords(fila, ['dni']);
  const nacimiento = getValByKeywords(fila, ['fecha de nacimiento', 'nacimiento']);
  const telefono = getValByKeywords(fila, ['teléfonos de contactos', 'teléfono', 'telefono', 'celular']);
  const lugar = getValByKeywords(fila, ['lugar']);
  const estadoCivil = getValByKeywords(fila, ['estado civil']);
  const fechaIngreso = getValByKeywords(fila, ['fecha de ing', 'ingreso']);
  const estudios = getValByKeywords(fila, ['estudios curs', 'estudios']);
  const titulo = getValByKeywords(fila, ['titulo obten', 'título']);
  const enfermedad = getValByKeywords(fila, ['enfermedad']);
  const lugarTrabajo = getValByKeywords(fila, ['lugar de trab', 'dependencia']);
  const actividad = getValByKeywords(fila, ['actividad', 'función']);

  setVal('val-fecha-realizacion', fechaRealizacion);
  setVal('val-legajo', legajo);
  setVal('val-nombre', nombre);
  setVal('val-domicilio', domicilio);
  setVal('val-dni', dni);
  setVal('val-nacimiento', formatearFecha(nacimiento));
  setVal('val-tel', telefono);
  setVal('val-lugar', lugar);
  setVal('val-estado-civil', estadoCivil);
  setVal('val-ingreso', formatearFecha(fechaIngreso));
  setVal('val-estudios', estudios);
  setVal('val-titulo', titulo);
  setVal('val-enfermedad', enfermedad);
  setVal('val-lugar-trabajo', lugarTrabajo);
  setVal('val-actividad', actividad);

  setVal('val-talle-camisa', getValByKeywords(fila, ['talle camisa']));
  setVal('val-talle-remera', getValByKeywords(fila, ['talle remera']));
  setVal('val-talle-pantalon', getValByKeywords(fila, ['talle pantal']));
  setVal('val-talle-calzado', getValByKeywords(fila, ['talle calzado']));

  renderizarRangoEnGrid('contenedor-conyuge', fila, col2idx('U'), col2idx('X'));
  renderizarHijosEstructurados('contenedor-hijos', fila);
  renderizarRangoEnGrid('contenedor-derechohabientes', fila, col2idx('BA'), col2idx('BI'));

  setVal('sig-nombre', nombre);
  setVal('sig-dni', dni);
  setVal('sig-fecha', fechaRealizacion);
  
  const padLegajo = String(legajo || '0000').padStart(4, '0');
  const padDni = String(dni || '0000').slice(-4);
  
  // Extraemos año actual o de la DDJJ para el ID de firma
  let anioFirma = '2026';
  if (fechaRealizacion && fechaRealizacion !== '-') {
    const partes = fechaRealizacion.split('/');
    if (partes.length >= 3) {
      anioFirma = partes[2].substring(0, 4);
    }
  }

  setVal('sig-id', `DDJJ-${anioFirma}-${padLegajo}-${padDni}`);
  setVal('sig-hash', `a7f98b${filaActual + 100}c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b`);
}

function renderizarHijosEstructurados(containerId, fila) {
  const container = document.getElementById(containerId);
  container.innerHTML = '';

  const bloquesHijos = [
    { titulo: "Hijo / Carga N° 1", inicio: col2idx('Y'), fin: col2idx('AE') },
    { titulo: "Hijo / Carga N° 2", inicio: col2idx('AF'), fin: col2idx('AL') },
    { titulo: "Hijo / Carga N° 3", inicio: col2idx('AM'), fin: col2idx('AS') },
    { titulo: "Hijo / Carga N° 4", inicio: col2idx('AT'), fin: col2idx('AZ') }
  ];

  let hayAlMenosUnHijo = false;
  let htmlGeneral = '';

  bloquesHijos.forEach(bloque => {
    let tieneDatosEsteHijo = false;
    let htmlCampos = '<div class="form-grid cols-3">';

    for (let i = bloque.inicio; i <= bloque.fin; i++) {
      const tituloCol = nombresColumnas[i];
      let valor = fila[i];

      if (valor !== undefined && valor !== null && String(valor).trim() !== "") {
        tieneDatosEsteHijo = true;
        hayAlMenosUnHijo = true;
        if (typeof valor === 'number' && valor > 30000 && valor < 60000) {
          valor = formatearFecha(valor);
        }
      } else {
        valor = "-";
      }

      htmlCampos += `
        <div class="field">
          <label>${tituloCol}</label>
          <div class="box">${valor}</div>
        </div>
      `;
    }

    htmlCampos += '</div>';

    if (tieneDatosEsteHijo) {
      htmlGeneral += `
        <div style="background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 4px; padding: 8px; margin-bottom: 8px;">
          <h4 style="margin: 0 0 6px 0; color: #1e3a8a; font-size: 11px; border-bottom: 1px solid #e2e8f0; padding-bottom: 2px;">
            👶 ${bloque.titulo}
          </h4>
          ${htmlCampos}
        </div>
      `;
    }
  });

  if (!hayAlMenosUnHijo) {
    container.innerHTML = '<div class="empty-section-msg">No registra hijos ni cargas de familia declaradas.</div>';
  } else {
    container.innerHTML = htmlGeneral;
  }
}

function renderizarRangoEnGrid(containerId, fila, idxInicio, idxFin) {
  const container = document.getElementById(containerId);
  container.innerHTML = '';

  let html = '<div class="form-grid cols-3">';
  let hayDatos = false;

  for (let i = idxInicio; i <= idxFin; i++) {
    const tituloColumna = nombresColumnas[i];
    let valor = fila[i];

    if (valor !== undefined && valor !== null && String(valor).trim() !== "") {
      hayDatos = true;
      if (typeof valor === 'number' && valor > 30000 && valor < 60000) {
        valor = formatearFecha(valor);
      }
    } else {
      valor = "-";
    }

    html += `
      <div class="field">
        <label>${tituloColumna}</label>
        <div class="box">${valor}</div>
      </div>
    `;
  }

  html += '</div>';

  if (!hayDatos) {
    container.innerHTML = '<div class="empty-section-msg">No registra datos declarados en este apartado.</div>';
  } else {
    container.innerHTML = html;
  }
}

function construirTablaModal() {
  const idxFecha = obtenerIndiceColumna(['registrado', 'marca']);
  const idxLegajo = obtenerIndiceColumna(['agente legajo', 'legajo']);
  const idxNombre = obtenerIndiceColumna(['nombre y apellido', 'agente']);
  const idxDni = obtenerIndiceColumna(['dni']);

  const indicesColumnas = [
    idxFecha !== -1 ? idxFecha : 0,
    idxLegajo !== -1 ? idxLegajo : 1,
    idxNombre !== -1 ? idxNombre : 2,
    idxDni !== -1 ? idxDni : 4
  ];

  const headerTr = document.getElementById('encabezadoTablaModal');
  const tbody = document.getElementById('cuerpoTablaModal');

  headerTr.innerHTML = '<th>#</th>';
  tbody.innerHTML = '';

  indicesColumnas.forEach(colIndex => {
    const th = document.createElement('th');
    th.textContent = nombresColumnas[colIndex] || `Columna ${colIndex + 1}`;
    headerTr.appendChild(th);
  });

  datosExcel.forEach((fila, index) => {
    const tr = document.createElement('tr');
    tr.onclick = () => {
      mostrarFila(index);
      cerrarModalListado();
    };

    let htmlRow = `<td><strong>${index + 1}</strong></td>`;
    
    indicesColumnas.forEach((colIndex, i) => {
      let valor = fila[colIndex];
      if (i === 0) valor = formatearFechaHora(valor);
      htmlRow += `<td>${(valor !== undefined && valor !== "") ? valor : '-'}</td>`;
    });

    tr.innerHTML = htmlRow;
    tbody.appendChild(tr);
  });
}

function abrirModalListado() { document.getElementById('modalListado').style.display = 'flex'; }
function cerrarModalListado() { document.getElementById('modalListado').style.display = 'none'; }

function imprimirTablaModal() {
  document.body.classList.add('printing-modal');
  window.print();
  document.body.classList.remove('printing-modal');
}

function setVal(id, valor) {
  const el = document.getElementById(id);
  if (el) el.textContent = (valor !== undefined && valor !== "") ? valor : "-";
}

function formatearFechaHora(val) {
  if (!val) return "-";
  if (typeof val === 'number') {
    const date = XLSX.SSF.parse_date_code(val);
    if(date) return `${date.d}/${date.m}/${date.y} ${date.H}:${date.M}:${date.S}`;
  }
  return String(val);
}

function formatearFecha(val) {
  if (!val) return "-";
  if (typeof val === 'number') {
    const date = XLSX.SSF.parse_date_code(val);
    if(date) return `${date.d}/${date.m}/${date.y}`;
  }
  return String(val);
}

function navegar(sentido) { mostrarFila(filaActual + sentido); }
function irAFila(valor) { if (valor !== "") mostrarFila(valor); }