var HOJAS = {
  MAESTROS:    "BASE MAESTROS",
  MATRIZ:      "MATRIZ RESUMEN",
  MAESTRO:     "Maestro",
  BITACORA:    "Bitácora",
  COMENTARIOS: "Comentarios"
};

// Columnas de hoja Maestro (base 0)
var COL_MAESTRO = {
  PROFESOR:      0,
  GRUPO:         1,
  ALUMNO:        2,
  CORREO:        3,
  FECHA:         4,
  ACTIVIDAD:     5,
  CALIFICACION:  6,
  MATERIA:       7
};

// Columnas de hoja BASE MAESTROS (base 0)
var COL_MAESTROS = { MATRICULA: 0, NOMBRE: 1, AREA: 2, CORREO: 3 };

// Columnas de hoja Comentarios (base 0)
var COL_COMENTARIOS = {
  FECHA:     0,
  PROFESOR:  1,
  CORREO:    2,
  ALUMNO:    3,
  MATERIA:   4,
  GRUPO:     5,
  TEXTO:     6
};

// Columnas de hoja Bitácora (base 0)
var COL_BITACORA = {
  FECHA:          0,
  PROFESOR:       1,
  GRUPOS_COUNT:   2,
  TIPO:           3,
  FILAS:          4,
  SEGUNDOS:       5,
  GRUPOS_DETALLE: 6   // JSON string con [{grupo, materia}]
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('WebApp')
    .setTitle('IBIME Calificaciones')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function normalizarNombre(str) {
  return String(str || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function extraerIdDeLink(link) {
  if (!link) return null;
  var m1 = link.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (m1) return m1[1];
  var m2 = link.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m2) return m2[1];
  return null;
}

// Asegura que una hoja exista; si no, la crea con encabezados
function asegurarHoja(ss, nombre, encabezados) {
  var h = ss.getSheetByName(nombre);
  if (!h) {
    h = ss.insertSheet(nombre);
    h.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
    h.getRange(1, 1, 1, encabezados.length)
      .setFontWeight("bold")
      .setBackground("#1a3a5c")
      .setFontColor("#ffffff");
    h.setFrozenRows(1);
  }
  return h;
}

function validarAdminPassword(password) {
  var correct = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || "ibime2025";
  return { ok: password === correct };
}

function validarLink(link) {
  try {
    var id = extraerIdDeLink(link);
    if (!id) return { ok: false, error: "Link inválido" };
    var ss = SpreadsheetApp.openById(id);
    return { ok: true, nombreHoja: ss.getName() };
  } catch (e) {
    return { ok: false, error: "Sin acceso: " + e.message };
  }
}

function obtenerCargaPorMatricula(matricula) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Buscar profesor en BASE MAESTROS
    var hMaestros = ss.getSheetByName(HOJAS.MAESTROS);
    if (!hMaestros) return { ok: false, error: "Hoja BASE MAESTROS no encontrada" };

    var dataMaestros = hMaestros.getDataRange().getValues();
    var profesor = null;
    for (var i = 1; i < dataMaestros.length; i++) {
      if (String(dataMaestros[i][COL_MAESTROS.MATRICULA]).trim() === String(matricula).trim()) {
        profesor = {
          matricula: String(dataMaestros[i][COL_MAESTROS.MATRICULA]).trim(),
          nombre:    String(dataMaestros[i][COL_MAESTROS.NOMBRE]).trim(),
          correo:    String(dataMaestros[i][COL_MAESTROS.CORREO]).trim()
        };
        break;
      }
    }
    if (!profesor) return { ok: false, error: "Matrícula no encontrada" };

    // 2. Leer MATRIZ RESUMEN para obtener carga asignada
    var hMatriz = ss.getSheetByName(HOJAS.MATRIZ);
    if (!hMatriz) return { ok: true, nombre: profesor.nombre, carga: [], gruposExistentes: [] };

    var dataMatriz = hMatriz.getRange(1, 1, hMatriz.getLastRow(), hMatriz.getLastColumn()).getValues();

    // Buscar columna del profesor (comparación normalizada)
    var colProf = -1;
    var nombreNorm = normalizarNombre(profesor.nombre);
    for (var c = 1; c < dataMatriz[0].length; c++) {
      if (normalizarNombre(dataMatriz[0][c]) === nombreNorm) {
        colProf = c;
        break;
      }
    }
    if (colProf === -1) return {
      ok:    true,
      nombre: profesor.nombre,
      carga:  [],
      gruposExistentes: [],
      aviso: "Profesor no encontrado en MATRIZ RESUMEN. Verifica que el nombre coincida."
    };

    // Recorrer filas de la Matriz para obtener grupos/materias asignados
    var carga = [];
    var grupoActual = "";
    for (var r = 1; r < dataMatriz.length; r++) {
      var g = String(dataMatriz[r][0]).trim();
      var materia = String(dataMatriz[r][colProf]).trim();
      if (g) grupoActual = g;
      if (!grupoActual || !materia) continue;

      var grupoUp = grupoActual.toUpperCase();
      var tipo = (
        grupoUp.includes("KEY") || grupoUp.includes("FCE") || grupoUp.includes("PET") ||
        grupoUp.includes("A1")  || grupoUp.includes("A2")  ||
        grupoUp.includes("B1")  || grupoUp.includes("B2")  || grupoUp.includes("C1")
      ) ? "Ingles" : "Español";

      carga.push({ grupo: grupoActual, materia: materia, tipo: tipo });
    }

    // 3. Leer hoja Maestro para saber qué grupos ya tiene cargados este profesor
    var gruposExistentes = [];
    var hMaestro = ss.getSheetByName(HOJAS.MAESTRO);
    if (hMaestro && hMaestro.getLastRow() > 1) {
      var dataMaestro = hMaestro.getDataRange().getValues();
      var vistos = {};
      for (var j = 1; j < dataMaestro.length; j++) {
        var profFila = normalizarNombre(String(dataMaestro[j][COL_MAESTRO.PROFESOR]).trim());
        if (profFila !== nombreNorm) continue;
        var grpF = String(dataMaestro[j][COL_MAESTRO.GRUPO]   ).trim();
        var matF = String(dataMaestro[j][COL_MAESTRO.MATERIA]  ).trim();
        var key  = grpF + "||" + matF;
        if (!vistos[key]) {
          vistos[key] = true;
          gruposExistentes.push({ grupo: grpF, materia: matF });
        }
      }
    }

    return {
      ok:                  true,
      nombre:              profesor.nombre,
      carga:               carga,
      gruposExistentes:    gruposExistentes,
      totalGruposAsignados: carga.length,
      tipoPredominante:    carga.length > 0 ? carga[0].tipo : ""
    };

  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function importarGruposWeb(datos) {
  var inicio = new Date();
  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var hMaestro  = ss.getSheetByName(HOJAS.MAESTRO);
    if (!hMaestro) return { ok: false, error: "Hoja 'Maestro' no encontrada" };

    // Asegurar que existan las hojas auxiliares
    asegurarHoja(ss, HOJAS.BITACORA, [
      "Fecha", "Profesor", "Grupos cargados", "Tipo", "Registros nuevos", "Segundos", "Detalle grupos (JSON)"
    ]);
    asegurarHoja(ss, HOJAS.COMENTARIOS, [
      "Fecha", "Profesor", "Correo alumno", "Alumno", "Materia", "Grupo", "Comentario"
    ]);

    // ── Construir SET de filas existentes en Maestro ──────────
    // Clave: normProf + "|" + grupo + "|" + materia + "|" + normAlumno + "|" + normActividad
    var existentes = {};
    var nombreNormProf = normalizarNombre(datos.profesor);

    if (hMaestro.getLastRow() > 1) {
      var dataMaestro = hMaestro.getDataRange().getValues();
      for (var j = 1; j < dataMaestro.length; j++) {
        var pNorm = normalizarNombre(String(dataMaestro[j][COL_MAESTRO.PROFESOR]).trim());
        if (pNorm !== nombreNormProf) continue;  // solo filas de este profesor (más rápido)
        var keyEx = [
          pNorm,
          String(dataMaestro[j][COL_MAESTRO.GRUPO]       ).trim().toLowerCase(),
          String(dataMaestro[j][COL_MAESTRO.MATERIA]     ).trim().toLowerCase(),
          normalizarNombre(String(dataMaestro[j][COL_MAESTRO.ALUMNO]   ).trim()),
          normalizarNombre(String(dataMaestro[j][COL_MAESTRO.ACTIVIDAD]).trim())
        ].join("|");
        existentes[keyEx] = true;
      }
    }

    // ── Procesar cada grupo enviado ───────────────────────────
    var totalNuevas = 0;
    var resumenLineas = [];
    var gruposImportados = [];
    var gruposDetalleBitacora = [];

    for (var g = 0; g < datos.gruposLinks.length; g++) {
      var item = datos.gruposLinks[g];
      var res  = procesarSheetClassroom(
        item.link, datos.profesor, item.grupo, item.asignatura,
        hMaestro, existentes, nombreNormProf
      );
      if (res.ok) {
        totalNuevas += res.filasNuevas;
        resumenLineas.push(
          "✅ " + item.grupo + " – " + item.asignatura + ": " +
          res.filasNuevas + " nuevas" +
          (res.filasOmitidas > 0 ? " (" + res.filasOmitidas + " duplicadas omitidas)" : "")
        );
        gruposImportados.push({ grupo: item.grupo, materia: item.asignatura });
        gruposDetalleBitacora.push({ grupo: item.grupo, materia: item.asignatura });
      } else {
        resumenLineas.push("❌ " + item.grupo + " – " + item.asignatura + ": " + res.error);
      }
    }

    var segundos = ((new Date() - inicio) / 1000).toFixed(1);

    // ── Registrar en Bitácora ─────────────────────────────────
    var hBitacora = ss.getSheetByName(HOJAS.BITACORA);
    hBitacora.appendRow([
      new Date(),
      datos.profesor,
      gruposDetalleBitacora.length,
      datos.tipoGrupo || "",
      totalNuevas,
      parseFloat(segundos),
      JSON.stringify(gruposDetalleBitacora)
    ]);

    return {
      ok:              true,
      mensajeFinal:    totalNuevas + " registros nuevos en " + segundos + "s",
      resumen:         resumenLineas.join("\n"),
      gruposImportados: gruposImportados
    };

  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function procesarSheetClassroom(link, profesor, grupo, asignatura, hMaestro, existentes, nombreNormProf) {
  try {
    var id = extraerIdDeLink(link);
    if (!id) return { ok: false, error: "Link inválido" };

    var data = SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getValues();
    if (data.length < 2) return { ok: false, error: "Sheet vacía" };

    // ── Detectar fila de actividades y fila de inicio de datos ─
    var filaActividades = 1;
    for (var r = 1; r < Math.min(data.length, 8); r++) {
      var cel = String(data[r][4] || "").trim();
      if (cel && cel !== "10" && isNaN(cel)) { filaActividades = r; break; }
    }

    var filaDatos = -1;
    for (var r2 = filaActividades + 1; r2 < data.length; r2++) {
      if (String(data[r2][2] || "").includes("@")) { filaDatos = r2; break; }
    }
    if (filaDatos === -1) return { ok: false, error: "No se encontraron alumnos" };

    // ── Leer actividades (columna 4 en adelante) ───────────────
    var actividades = [];
    for (var c = 4; c < data[filaActividades].length; c++) {
      var nom = String(data[filaActividades][c] || "").trim();
      if (!nom) continue;
      actividades.push({
        col:    c,
        nombre: nom,
        fecha:  String(data[0][c] || "Sin fecha").trim()
      });
    }
    if (actividades.length === 0) return { ok: false, error: "Sin actividades detectadas" };

    // ── Construir filas nuevas verificando duplicados ──────────
    var nuevasFilas = [];
    var filasOmitidas = 0;
    var grupoLower     = grupo.trim().toLowerCase();
    var asignaturaLow  = asignatura.trim().toLowerCase();

    for (var ra = filaDatos; ra < data.length; ra++) {
      var correo = String(data[ra][2] || "").trim();
      if (!correo || !correo.includes("@")) continue;

      var nombreCompleto = (String(data[ra][0] || "") + " " + String(data[ra][1] || "")).trim();
      var alumnoNorm     = normalizarNombre(nombreCompleto);

      for (var a = 0; a < actividades.length; a++) {
        var act = actividades[a];
        var actNorm = normalizarNombre(act.nombre);

        // Verificar duplicado
        var keyCheck = [
          nombreNormProf,
          grupoLower,
          asignaturaLow,
          alumnoNorm,
          actNorm
        ].join("|");

        if (existentes[keyCheck]) {
          filasOmitidas++;
          continue;
        }

        // Parsear calificación
        var val   = data[ra][act.col];
        var calif = "";
        if (val !== "" && val !== null && val !== undefined) {
          var num = parseFloat(String(val).replace(",", "."));
          calif = isNaN(num) ? String(val) : num;
        }

        nuevasFilas.push([
          profesor,        // COL 0 Profesor
          grupo,           // COL 1 Grupo
          nombreCompleto,  // COL 2 Alumno
          correo,          // COL 3 Correo
          act.fecha,       // COL 4 Fecha
          act.nombre,      // COL 5 Actividad
          calif,           // COL 6 Calificación
          asignatura       // COL 7 Materia (Asignatura)
        ]);

        // Agregar al set para evitar duplicados dentro del mismo lote
        existentes[keyCheck] = true;
      }
    }

    if (nuevasFilas.length > 0) {
      hMaestro.getRange(hMaestro.getLastRow() + 1, 1, nuevasFilas.length, 8).setValues(nuevasFilas);
    }

    return {
      ok:           true,
      filasNuevas:  nuevasFilas.length,
      filasOmitidas: filasOmitidas
    };

  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function obtenerReporteProfesor(matricula) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Buscar profesor
    var dataMaestros = ss.getSheetByName(HOJAS.MAESTROS).getDataRange().getValues();
    var nombreProf = null;
    for (var i = 1; i < dataMaestros.length; i++) {
      if (String(dataMaestros[i][COL_MAESTROS.MATRICULA]).trim() === String(matricula).trim()) {
        nombreProf = String(dataMaestros[i][COL_MAESTROS.NOMBRE]).trim();
        break;
      }
    }
    if (!nombreProf) return { ok: false, error: "Matrícula no encontrada" };

    var nombreNorm = normalizarNombre(nombreProf);

    // Grupos cargados en Maestro
    var hMaestro = ss.getSheetByName(HOJAS.MAESTRO);
    var grupos = [];
    if (hMaestro && hMaestro.getLastRow() > 1) {
      var data = hMaestro.getDataRange().getValues();
      var vistos = {};
      for (var j = 1; j < data.length; j++) {
        if (normalizarNombre(String(data[j][COL_MAESTRO.PROFESOR]).trim()) !== nombreNorm) continue;
        var grp = String(data[j][COL_MAESTRO.GRUPO]  ).trim();
        var mat = String(data[j][COL_MAESTRO.MATERIA] ).trim();
        if (!grp || !mat) continue;
        var key = grp + "||" + mat;
        if (!vistos[key]) { vistos[key] = true; grupos.push({ grupo: grp, materia: mat }); }
      }
    }

    // Total asignado en MATRIZ RESUMEN
    var totalAsignados = 0;
    var hMatriz = ss.getSheetByName(HOJAS.MATRIZ);
    if (hMatriz) {
      var dataMatriz = hMatriz.getRange(1, 1, hMatriz.getLastRow(), hMatriz.getLastColumn()).getValues();
      var colProf = -1;
      for (var c = 1; c < dataMatriz[0].length; c++) {
        if (normalizarNombre(String(dataMatriz[0][c]).trim()) === nombreNorm) { colProf = c; break; }
      }
      if (colProf !== -1) {
        for (var r = 1; r < dataMatriz.length; r++) {
          if (String(dataMatriz[r][colProf]).trim()) totalAsignados++;
        }
      }
    }

    return {
      ok:                   true,
      nombre:               nombreProf,
      grupos:               grupos,
      totalGruposAsignados: totalAsignados || grupos.length
    };

  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function obtenerAlumnosPorGrupoMateria(matricula, grupo, materia) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Buscar nombre del profesor
    var dataMaestros = ss.getSheetByName(HOJAS.MAESTROS).getDataRange().getValues();
    var nombreProf = null;
    for (var i = 1; i < dataMaestros.length; i++) {
      if (String(dataMaestros[i][COL_MAESTROS.MATRICULA]).trim() === String(matricula).trim()) {
        nombreProf = String(dataMaestros[i][COL_MAESTROS.NOMBRE]).trim();
        break;
      }
    }
    if (!nombreProf) return { ok: false, error: "Profesor no encontrado" };

    var hMaestro = ss.getSheetByName(HOJAS.MAESTRO);
    if (!hMaestro || hMaestro.getLastRow() < 2) return { ok: true, alumnos: [] };

    var data = hMaestro.getDataRange().getValues();
    var nombreNorm = normalizarNombre(nombreProf);

    // Mapa de alumnos: correo → { nombre, correo, actividades[] }
    var alumnos = {};
    for (var j = 1; j < data.length; j++) {
      if (normalizarNombre(String(data[j][COL_MAESTRO.PROFESOR]).trim()) !== nombreNorm) continue;
      if (String(data[j][COL_MAESTRO.GRUPO]  ).trim() !== grupo)   continue;
      if (String(data[j][COL_MAESTRO.MATERIA]).trim() !== materia)  continue;

      var correo = String(data[j][COL_MAESTRO.CORREO]).trim();
      if (!correo) continue;

      if (!alumnos[correo]) {
        alumnos[correo] = {
          nombre:             String(data[j][COL_MAESTRO.ALUMNO]).trim(),
          correo:             correo,
          actividadesDetalle: [],
          califs:             []
        };
      }

      var actNombre = String(data[j][COL_MAESTRO.ACTIVIDAD    ]).trim();
      var fecha     = String(data[j][COL_MAESTRO.FECHA        ]).trim();
      var califRaw  = data[j][COL_MAESTRO.CALIFICACION];
      var califNum  = parseFloat(String(califRaw).replace(",", "."));
      var califFinal = isNaN(califNum) ? String(califRaw || "") : califNum;

      alumnos[correo].actividadesDetalle.push({
        nombre:        actNombre,
        fecha:         fecha,
        calificacion:  califFinal
      });

      if (!isNaN(califNum)) alumnos[correo].califs.push(califNum);
    }

    // Leer comentarios de la hoja COMENTARIOS
    var comentariosMapa = {};
    var hCom = ss.getSheetByName(HOJAS.COMENTARIOS);
    if (hCom && hCom.getLastRow() > 1) {
      var dataCom = hCom.getDataRange().getValues();
      for (var k = 1; k < dataCom.length; k++) {
        var correoC  = String(dataCom[k][COL_COMENTARIOS.CORREO]  ).trim();
        var materiaC = String(dataCom[k][COL_COMENTARIOS.MATERIA] ).trim();
        var grupoC   = String(dataCom[k][COL_COMENTARIOS.GRUPO]   ).trim();
        if (correoC && materiaC === materia && grupoC === grupo) {
          comentariosMapa[correoC] = String(dataCom[k][COL_COMENTARIOS.TEXTO]).trim();
        }
      }
    }

    // Armar resultado final
    var resultado = [];
    for (var correoKey in alumnos) {
      var al   = alumnos[correoKey];
      var prom = al.califs.length > 0
        ? Math.round(al.califs.reduce(function(a, b) { return a + b; }, 0) / al.califs.length * 10) / 10
        : null;

      resultado.push({
        nombre:             al.nombre,
        correo:             al.correo,
        promedio:           prom,
        actividades:        al.actividadesDetalle.length,
        actividadesDetalle: al.actividadesDetalle,
        comentario:         comentariosMapa[correoKey] || ""
      });
    }

    // Ordenar por nombre
    resultado.sort(function(a, b) { return a.nombre.localeCompare(b.nombre, 'es'); });

    return { ok: true, alumnos: resultado };

  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function obtenerActividadesAlumno(datos) {
  try {
    var correo  = String(datos.correo  || "").trim();
    var materia = String(datos.materia || "").trim();
    var grupo   = String(datos.grupo   || "").trim();
    if (!correo) return { ok: false, actividades: [] };

    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var hMaestro = ss.getSheetByName(HOJAS.MAESTRO);
    if (!hMaestro || hMaestro.getLastRow() < 2) return { ok: true, actividades: [] };

    var data = hMaestro.getDataRange().getValues();
    var actividades = [];

    for (var j = 1; j < data.length; j++) {
      if (String(data[j][COL_MAESTRO.CORREO ]).trim() !== correo)   continue;
      if (materia && String(data[j][COL_MAESTRO.MATERIA]).trim() !== materia) continue;
      if (grupo   && String(data[j][COL_MAESTRO.GRUPO  ]).trim() !== grupo)   continue;

      var califRaw = data[j][COL_MAESTRO.CALIFICACION];
      var califNum = parseFloat(String(califRaw).replace(",", "."));

      actividades.push({
        nombre:       String(data[j][COL_MAESTRO.ACTIVIDAD]).trim(),
        fecha:        String(data[j][COL_MAESTRO.FECHA    ]).trim(),
        calificacion: isNaN(califNum) ? String(califRaw || "") : califNum
      });
    }

    return { ok: true, actividades: actividades };

  } catch (e) {
    return { ok: false, error: e.message, actividades: [] };
  }
}

function guardarComentario(datos) {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var hCom = asegurarHoja(ss, HOJAS.COMENTARIOS, [
      "Fecha", "Profesor", "Correo alumno", "Alumno", "Materia", "Grupo", "Comentario"
    ]);

    var correo  = String(datos.correo   || "").trim();
    var materia = String(datos.materia  || "").trim();
    var grupo   = String(datos.grupo    || "").trim();
    var texto   = String(datos.texto    || "").trim();
    var profesor= String(datos.profesor || "").trim();
    var alumno  = String(datos.alumno   || "").trim();

    // Buscar si ya existe fila para este correo + materia + grupo
    var filaExistente = -1;
    if (hCom.getLastRow() > 1) {
      var dataCom = hCom.getDataRange().getValues();
      for (var i = 1; i < dataCom.length; i++) {
        if (
          String(dataCom[i][COL_COMENTARIOS.CORREO  ]).trim() === correo  &&
          String(dataCom[i][COL_COMENTARIOS.MATERIA ]).trim() === materia &&
          String(dataCom[i][COL_COMENTARIOS.GRUPO   ]).trim() === grupo
        ) {
          filaExistente = i + 1; // número de fila en la hoja (base 1)
          break;
        }
      }
    }

    var ahora = new Date();

    if (filaExistente !== -1) {
      // Actualizar fila existente
      hCom.getRange(filaExistente, COL_COMENTARIOS.FECHA    + 1).setValue(ahora);
      hCom.getRange(filaExistente, COL_COMENTARIOS.PROFESOR + 1).setValue(profesor);
      hCom.getRange(filaExistente, COL_COMENTARIOS.TEXTO    + 1).setValue(texto);
    } else {
      // Agregar nueva fila
      hCom.appendRow([ahora, profesor, correo, alumno, materia, grupo, texto]);
    }

    return { ok: true };

  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function obtenerDatosAdmin() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var totalCalificaciones = 0;

    var hMaestro = ss.getSheetByName(HOJAS.MAESTRO);
    var mapaTablero = {};
    var entregadosMapa = {};

    if (hMaestro && hMaestro.getLastRow() > 1) {
      var dM = hMaestro.getDataRange().getValues();
      for (var j = 1; j < dM.length; j++) {
        var prof = String(dM[j][COL_MAESTRO.PROFESOR] || "").trim();
        var grp  = String(dM[j][COL_MAESTRO.GRUPO]    || "").trim();
        var mat  = String(dM[j][COL_MAESTRO.MATERIA]  || "").trim();
        if (!prof) continue;
        totalCalificaciones++;
        var pNorm = normalizarNombre(prof);
        if (!mapaTablero[pNorm]) mapaTablero[pNorm] = { profOriginal: prof, cargados: {} };
        if (grp && mat) mapaTablero[pNorm].cargados[grp + "||" + mat] = { grupo: grp, materia: mat };
        if (!entregadosMapa[pNorm]) entregadosMapa[pNorm] = {};
        if (grp && mat) entregadosMapa[pNorm][grp + "||" + mat] = true;
      }
    }

    var asignadosMapa = {};
    var profMatriz = {};
    var hMatriz = ss.getSheetByName(HOJAS.MATRIZ);

    if (hMatriz && hMatriz.getLastRow() > 1) {
      var dMz = hMatriz.getRange(1, 1, hMatriz.getLastRow(), hMatriz.getLastColumn()).getValues();
      for (var nc = 1; nc < dMz[0].length; nc++) {
        var np = String(dMz[0][nc] || "").trim();
        if (!np) continue;
        var npNorm = normalizarNombre(np);
        if (!asignadosMapa[npNorm]) asignadosMapa[npNorm] = [];
        if (!profMatriz[npNorm]) profMatriz[npNorm] = { nombre: np, esperados: 0 };
        var grupoActual2 = "";
        for (var nr = 1; nr < dMz.length; nr++) {
          var gCell = String(dMz[nr][0] || "").trim();
          var mCell = String(dMz[nr][nc] || "").trim();
          if (gCell) grupoActual2 = gCell;
          if (grupoActual2 && mCell) {
            asignadosMapa[npNorm].push({ grupo: grupoActual2, materia: mCell });
            profMatriz[npNorm].esperados++;
          }
        }
      }
    }

    var tablero = [];
    var todosProfs = {};
    for (var pn in mapaTablero)   todosProfs[pn] = mapaTablero[pn].profOriginal;
    for (var an in asignadosMapa) if (!todosProfs[an]) todosProfs[an] = an;

    for (var tn in todosProfs) {
      var cargadosObj  = mapaTablero[tn] ? mapaTablero[tn].cargados : {};
      var asignadosArr = asignadosMapa[tn] ? asignadosMapa[tn] : [];
      var cargadosArr  = [];
      for (var ck in cargadosObj) { cargadosArr.push(cargadosObj[ck]); }
      var faltantesArr = [];
      for (var fa = 0; fa < asignadosArr.length; fa++) {
        var ag = asignadosArr[fa];
        if (!cargadosObj[ag.grupo + "||" + ag.materia]) faltantesArr.push(ag);
      }
      tablero.push({
        profesor:       todosProfs[tn],
        cargados:       cargadosArr,
        faltantes:      faltantesArr,
        totalAsignados: asignadosArr.length,
        totalCargados:  cargadosArr.length
      });
    }
    tablero.sort(function(a, b) { return b.faltantes.length - a.faltantes.length; });
    var bitacora = [];
    var hBit = ss.getSheetByName(HOJAS.BITACORA);
    if (hBit && hBit.getLastRow() > 1) {
      var dBit = hBit.getDataRange().getValues();
      var inicioB = Math.max(1, dBit.length - 50);
      for (var bi = dBit.length - 1; bi >= inicioB; bi--) {
        var gdRaw = String(dBit[bi][COL_BITACORA.GRUPOS_DETALLE] || "").trim();
        var gdParsed = [];
        if (gdRaw) { try { gdParsed = JSON.parse(gdRaw); } catch(ep) {} }
        bitacora.push({
          fecha:         dBit[bi][COL_BITACORA.FECHA],
          profesor:      String(dBit[bi][COL_BITACORA.PROFESOR]    || "").trim(),
          grupos:        dBit[bi][COL_BITACORA.GRUPOS_COUNT] || 0,
          tipo:          String(dBit[bi][COL_BITACORA.TIPO]        || "").trim(),
          filas:         dBit[bi][COL_BITACORA.FILAS]        || 0,
          segundos:      dBit[bi][COL_BITACORA.SEGUNDOS]     || 0,
          gruposDetalle: gdParsed
        });
      }
    }

    return {
      ok:                  true,
      totalCalificaciones: totalCalificaciones,
      tablero:             tablero,
      bitacora:            bitacora
    };

  } catch (e) {
    return { ok: false, error: "obtenerDatosAdmin: " + e.message };
  }
}

function repararNombresEnHojaMaestro() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hMaestros = ss.getSheetByName(HOJAS.MAESTROS);
  if (!hMaestros) { Logger.log("❌ Hoja BASE MAESTROS no encontrada"); return; }

  var dataMaestros = hMaestros.getDataRange().getValues();
  var canonicos = {};
  for (var i = 1; i < dataMaestros.length; i++) {
    var nombre = String(dataMaestros[i][COL_MAESTROS.NOMBRE] || "").trim();
    if (nombre) canonicos[normalizarNombre(nombre)] = nombre;
  }

  var hMaestro = ss.getSheetByName(HOJAS.MAESTRO);
  if (!hMaestro || hMaestro.getLastRow() < 2) { Logger.log("⚠️ Hoja Maestro vacía"); return; }

  var data = hMaestro.getDataRange().getValues();
  var corregidos = 0, noEncontrados = [];

  for (var j = 1; j < data.length; j++) {
    var actual   = String(data[j][COL_MAESTRO.PROFESOR] || "").trim();
    if (!actual) continue;
    var canonico = canonicos[normalizarNombre(actual)];
    if (canonico && canonico !== actual) {
      hMaestro.getRange(j + 1, COL_MAESTRO.PROFESOR + 1).setValue(canonico);
      corregidos++;
    } else if (!canonico) {
      if (noEncontrados.indexOf(actual) === -1) noEncontrados.push(actual);
    }
  }

  Logger.log("✅ " + corregidos + " celdas corregidas");
  if (noEncontrados.length > 0) Logger.log("⚠️ Sin match: " + noEncontrados.join(", "));

  SpreadsheetApp.getUi().alert(
    "Reparación completa\n\nCeldas corregidas: " + corregidos + "\n" +
    (noEncontrados.length > 0 ? "Sin match: " + noEncontrados.join(", ") : "Todos los nombres encontrados ✅")
  );
}

function generarReporteAlumnosSinCorreo() {
  try {
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var hMaestro = ss.getSheetByName(HOJAS.MAESTRO);
    if (!hMaestro || hMaestro.getLastRow() < 2) return { ok: false, error: "Hoja Maestro vacía" };

    var data = hMaestro.getDataRange().getValues();
    var sinCorreo = [];
    for (var i = 1; i < data.length; i++) {
      var alumno = String(data[i][COL_MAESTRO.ALUMNO ] || "").trim();
      var correo = String(data[i][COL_MAESTRO.CORREO ] || "").trim();
      if (!alumno || (correo && correo.includes("@"))) continue;
      sinCorreo.push([
        String(data[i][COL_MAESTRO.PROFESOR]).trim(),
        String(data[i][COL_MAESTRO.GRUPO   ]).trim(),
        String(data[i][COL_MAESTRO.MATERIA ]).trim(),
        alumno,
        correo || "(vacío)",
        i + 1
      ]);
    }

    var NOMBRE_HOJA = "Alumnos Sin Correo";
    var hSC = ss.getSheetByName(NOMBRE_HOJA) || ss.insertSheet(NOMBRE_HOJA);
    hSC.clearContents(); hSC.clearFormats();
    hSC.getRange(1,1,1,6).setValues([["Profesor","Grupo","Materia","Alumno","Correo actual","Fila en Maestro"]]);
    hSC.getRange(1,1,1,6).setFontWeight("bold").setBackground("#1a3a5c").setFontColor("#ffffff");
    if (sinCorreo.length > 0) hSC.getRange(2,1,sinCorreo.length,6).setValues(sinCorreo);
    hSC.setFrozenRows(1);

    return { ok: true, total: sinCorreo.length };
  } catch (e) { return { ok: false, error: e.message }; }
}
