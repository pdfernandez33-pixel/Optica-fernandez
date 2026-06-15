// ID DE TU GOOGLE SHEET - REEMPLAZA CON EL TUYO
const SHEET_ID = '1kc7lASHrajsACW6YeFFpcSpxYBbVqBgq6XNcrNwoGpE';

// ========================================
// FUNCIÓN PRINCIPAL - Llamada desde HTML
// ========================================
function doGet(e) {
  return HtmlService.createHtmlOutput('Google Apps Script está funcionando');
}

// ========================================
// OBTENER DATOS DEL DASHBOARD
// ========================================
function obtenerDashboard() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    // Obtener datos de cada hoja
    const ventasHoy = contarVentasHoy(ss);
    const totalPacientes = contarPacientes(ss);
    const ingresosDelMes = calcularIngresosDelMes(ss);
    const alertasStock = contarAlertasStock(ss);
    const ultimasVentas = obtenerUltimasVentas(ss, 5);
    
    return {
      ventasHoy: ventasHoy,
      pacientes: totalPacientes,
      ingresos: ingresosDelMes.toFixed(2),
      alertas: alertasStock,
      ultimasVentas: ultimasVentas,
      error: null
    };
  } catch (error) {
    return {
      error: 'Error al obtener datos: ' + error.toString()
    };
  }
}

// ========================================
// CONTAR VENTAS DE HOY
// ========================================
function contarVentasHoy(ss) {
  try {
    const ventasSheet = ss.getSheetByName('VENTAS');
    if (!ventasSheet) return 0;
    
    const datos = ventasSheet.getDataRange().getValues();
    const hoy = new Date();
    const fechaHoy = (hoy.getDate()).toString().padStart(2, '0') + '/' + 
                     (hoy.getMonth() + 1).toString().padStart(2, '0') + '/' + 
                     hoy.getFullYear();
    
    let contador = 0;
    for (let i = 1; i < datos.length; i++) {
      const fechaVenta = datos[i][0] ? datos[i][0].toString() : '';
      if (fechaVenta.includes(fechaHoy.slice(0, 10))) {
        contador++;
      }
    }
    return contador;
  } catch (error) {
    Logger.log('Error en contarVentasHoy: ' + error);
    return 0;
  }
}

// ========================================
// CONTAR TOTAL DE PACIENTES
// ========================================
function contarPacientes(ss) {
  try {
    const pacientesSheet = ss.getSheetByName('PACIENTES');
    if (!pacientesSheet) return 0;
    
    // Contar filas que tengan datos (excluyendo encabezado)
    const datos = pacientesSheet.getDataRange().getValues();
    return Math.max(0, datos.length - 1);
  } catch (error) {
    Logger.log('Error en contarPacientes: ' + error);
    return 0;
  }
}

// ========================================
// CALCULAR INGRESOS DEL MES
// ========================================
function calcularIngresosDelMes(ss) {
  try {
    const ventasSheet = ss.getSheetByName('VENTAS');
    if (!ventasSheet) return 0;
    
    const datos = ventasSheet.getDataRange().getValues();
    const hoy = new Date();
    const mesActual = hoy.getMonth() + 1;
    const anioActual = hoy.getFullYear();
    
    let total = 0;
    
    // Columna F es TOTAL (índice 5)
    for (let i = 1; i < datos.length; i++) {
      const fechaStr = datos[i][0] ? datos[i][0].toString() : '';
      const totalVenta = datos[i][5] || 0;
      
      // Verificar si la fecha es del mes actual
      if (fechaStr.includes('/' + mesActual.toString().padStart(2, '0'))) {
        total += parseFloat(totalVenta) || 0;
      }
    }
    
    return total;
  } catch (error) {
    Logger.log('Error en calcularIngresosDelMes: ' + error);
    return 0;
  }
}

// ========================================
// CONTAR ALERTAS DE STOCK
// ========================================
function contarAlertasStock(ss) {
  try {
    const inventarioSheet = ss.getSheetByName('INVENTARIO');
    if (!inventarioSheet) return 0;
    
    const datos = inventarioSheet.getDataRange().getValues();
    let alertas = 0;
    
    // Verificar si hay columna de stock mínimo (ajusta según tu estructura)
    for (let i = 1; i < datos.length; i++) {
      const stock = parseFloat(datos[i][2]) || 0;
      const stockMinimo = parseFloat(datos[i][3]) || 0;
      
      if (stock <= stockMinimo && stock > 0) {
        alertas++;
      }
    }
    
    return alertas;
  } catch (error) {
    Logger.log('Error en contarAlertasStock: ' + error);
    return 0;
  }
}

// ========================================
// OBTENER ÚLTIMAS VENTAS
// ========================================
function obtenerUltimasVentas(ss, limite = 5) {
  try {
    const ventasSheet = ss.getSheetByName('VENTAS');
    if (!ventasSheet) return [];
    
    const datos = ventasSheet.getDataRange().getValues();
    const ventas = [];
    
    // Recorrer de atrás hacia adelante para obtener las últimas
    for (let i = Math.min(datos.length - 1, datos.length - 1); i >= Math.max(1, datos.length - limite - 1); i--) {
      if (datos[i][0]) { // Si hay fecha
        ventas.push({
          fecha: datos[i][0].toString().slice(0, 10),
          cliente: datos[i][1] || 'N/A',
          dni: datos[i][2] || 'N/A',
          productos: datos[i][4] || 'N/A',
          total: '$' + (parseFloat(datos[i][5]) || 0).toFixed(2),
          vendedor: datos[i][7] || 'N/A'
        });
      }
    }
    
    return ventas;
  } catch (error) {
    Logger.log('Error en obtenerUltimasVentas: ' + error);
    return [];
  }
}

// ========================================
// OBTENER PACIENTES RECIENTES
// ========================================
function obtenerPacientesRecientes(limite = 5) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const pacientesSheet = ss.getSheetByName('PACIENTES');
    if (!pacientesSheet) return [];
    
    const datos = pacientesSheet.getDataRange().getValues();
    const pacientes = [];
    
    for (let i = Math.min(datos.length - 1, datos.length - 1); i >= Math.max(1, datos.length - limite - 1); i--) {
      if (datos[i][1]) { // Si hay nombre
        pacientes.push({
          fecha: datos[i][0].toString().slice(0, 10),
          nombre: datos[i][1] || 'N/A',
          telefono: datos[i][3] || 'N/A',
          estado: datos[i][7] || 'Activo'
        });
      }
    }
    
    return pacientes;
  } catch (error) {
    Logger.log('Error en obtenerPacientesRecientes: ' + error);
    return [];
  }
}

// ========================================
// AGREGAR NUEVA VENTA
// ========================================
function agregarVenta(fecha, cliente, dni, telefono, productos, total, formaPago, vendedor, sede) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const ventasSheet = ss.getSheetByName('VENTAS');
    if (!ventasSheet) return { error: 'Hoja VENTAS no encontrada' };
    
    ventasSheet.appendRow([fecha, cliente, dni, telefono, productos, total, formaPago, vendedor, sede]);
    
    return { success: true, message: 'Venta agregada correctamente' };
  } catch (error) {
    return { error: 'Error al agregar venta: ' + error.toString() };
  }
}

// ========================================
// AGREGAR NUEVO PACIENTE
// ========================================
function agregarPaciente(fecha, nombre, whatsapp, email, ultimaAtencion, proximaRevision, estado, observaciones, recordatorio, dni) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const pacientesSheet = ss.getSheetByName('PACIENTES');
    if (!pacientesSheet) return { error: 'Hoja PACIENTES no encontrada' };
    
    pacientesSheet.appendRow([fecha, nombre, whatsapp, email, ultimaAtencion, proximaRevision, estado, observaciones, recordatorio, dni]);
    
    return { success: true, message: 'Paciente agregado correctamente' };
  } catch (error) {
    return { error: 'Error al agregar paciente: ' + error.toString() };
  }
}

// ========================================
// AGREGAR NUEVA RECETA
// ========================================
function agregarReceta(fecha, pacienteId, pacienteNombre, odEsfera, odCil, odEje, odAdd, oiEsfera, oiCil) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const recetasSheet = ss.getSheetByName('RECETAS');
    if (!recetasSheet) return { error: 'Hoja RECETAS no encontrada' };
    
    recetasSheet.appendRow([fecha, pacienteId, pacienteNombre, odEsfera, odCil, odEje, odAdd, oiEsfera, oiCil]);
    
    return { success: true, message: 'Receta agregada correctamente' };
  } catch (error) {
    return { error: 'Error al agregar receta: ' + error.toString() };
  }
}

// ========================================
// BÚSQUEDA DE PACIENTES
// ========================================
function buscarPaciente(termino) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const pacientesSheet = ss.getSheetByName('PACIENTES');
    if (!pacientesSheet) return [];
    
    const datos = pacientesSheet.getDataRange().getValues();
    const resultados = [];
    const terminoLower = termino.toLowerCase();
    
    for (let i = 1; i < datos.length; i++) {
      const nombre = (datos[i][1] || '').toLowerCase();
      const telefono = (datos[i][2] || '').toLowerCase();
      const email = (datos[i][3] || '').toLowerCase();
      
      if (nombre.includes(terminoLower) || telefono.includes(terminoLower) || email.includes(terminoLower)) {
        resultados.push({
          id: i,
          nombre: datos[i][1],
          telefono: datos[i][2],
          email: datos[i][3],
          estado: datos[i][7]
        });
      }
    }
    
    return resultados;
  } catch (error) {
    Logger.log('Error en buscarPaciente: ' + error);
    return [];
  }
}

// ========================================
// OBTENER INVENTARIO
// ========================================
function obtenerInventario() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const inventarioSheet = ss.getSheetByName('INVENTARIO');
    if (!inventarioSheet) return [];
    
    const datos = inventarioSheet.getDataRange().getValues();
    const inventario = [];
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0]) { // Si hay nombre de producto
        inventario.push({
          nombre: datos[i][0] || 'N/A',
          stock: datos[i][2] || 0,
          stockMinimo: datos[i][3] || 0,
          precio: datos[i][4] || 0
        });
      }
    }
    
    return inventario;
  } catch (error) {
    Logger.log('Error en obtenerInventario: ' + error);
    return [];
  }
}

// ========================================
// OBTENER DEUDAS PENDIENTES
// ========================================
function obtenerDeudas() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const deudasSheet = ss.getSheetByName('DEUDAS');
    if (!deudasSheet) return [];
    
    const datos = deudasSheet.getDataRange().getValues();
    const deudas = [];
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0]) { // Si hay fecha
        deudas.push({
          fecha: datos[i][0].toString().slice(0, 10),
          cliente: datos[i][1] || 'N/A',
          monto: '$' + (parseFloat(datos[i][2]) || 0).toFixed(2),
          estado: datos[i][3] || 'Pendiente'
        });
      }
    }
    
    return deudas;
  } catch (error) {
    Logger.log('Error en obtenerDeudas: ' + error);
    return [];
  }
}
