// Guardar contexto global
let excelContext = null;

// Espera a que Office.js esté listo
Office.onReady((reason) => {
    console.log('Office.onReady ejecutado:', reason);
    
    if (reason === Office.HostType.Excel) {
        console.log('✅ Excel Web Add-in iniciado correctamente');
        inicializarComplemento();
    } else {
        console.error('❌ No es Excel');
    }
});

/**
 * Inicializar el complemento
 */
function inicializarComplemento() {
    console.log('Inicializando complemento...');
    
    try {
        // Actualizar la hora cada segundo
        actualizarHoraDisplay();
        setInterval(actualizarHoraDisplay, 1000);
        
        // Asignar eventos a los botones
        const btnHora = document.getElementById('btnInsertarHora');
        const btnFecha = document.getElementById('btnInsertarFecha');
        
        if (btnHora) {
            btnHora.addEventListener('click', insertarHora);
            console.log('✅ Botón de hora asignado');
        } else {
            console.error('❌ No se encontró btnInsertarHora');
        }
        
        if (btnFecha) {
            btnFecha.addEventListener('click', insertarFechaHora);
            console.log('✅ Botón de fecha asignado');
        } else {
            console.error('❌ No se encontró btnInsertarFecha');
        }
        
        console.log('✅ Complemento inicializado correctamente');
    } catch (error) {
        console.error('❌ Error en inicializarComplemento:', error);
        mostrarEstatus('Error al inicializar: ' + error.message, 'error');
    }
}

/**
 * Actualiza la pantalla de la hora actual
 */
function actualizarHoraDisplay() {
    try {
        const ahora = new Date();
        const horas = String(ahora.getHours()).padStart(2, '0');
        const minutos = String(ahora.getMinutes()).padStart(2, '0');
        const segundos = String(ahora.getSeconds()).padStart(2, '0');
        
        const horaFormato = `${horas}:${minutos}:${segundos}`;
        const elemento = document.getElementById('horaActual');
        
        if (elemento) {
            elemento.textContent = horaFormato;
        } else {
            console.error('❌ No se encontró elemento horaActual');
        }
    } catch (error) {
        console.error('❌ Error en actualizarHoraDisplay:', error);
    }
}

/**
 * Inserta la hora en la celda activa
 */
async function insertarHora() {
    console.log('Insertando hora...');
    
    try {
        const formato = document.querySelector('input[name="formato"]:checked');
        
        if (!formato) {
            mostrarEstatus('❌ Selecciona un formato', 'error');
            return;
        }
        
        const tipoFormato = formato.value;
        console.log('Formato seleccionado:', tipoFormato);
        
        await Excel.run(async (context) => {
            console.log('Excel.run iniciado');
            
            const cell = context.application.activeCell;
            cell.load('address');
            
            await context.sync();
            console.log('Celda activa:', cell.address);
            
            // Obtener la hora actual
            const ahora = new Date();
            let valor;
            
            switch(tipoFormato) {
                case 'hora':
                    valor = formatearHora(ahora);
                    break;
                case 'horaAmPm':
                    valor = formatearHoraAmPm(ahora);
                    break;
                case 'horaMinutos':
                    valor = formatearMinutosSegundos(ahora);
                    break;
                case 'fechaHora':
                    valor = formatearFechaHora(ahora);
                    break;
                default:
                    valor = formatearHora(ahora);
            }
            
            console.log('Valor a insertar:', valor);
            
            // Insertar valor en la celda activa
            cell.values = [[valor]];
            
            await context.sync();
            
            console.log('✅ Hora insertada exitosamente');
            mostrarEstatus(`✅ Hora insertada en ${cell.address}: ${valor}`, 'success');
        });
    } catch (error) {
        console.error('❌ Error al insertar hora:', error);
        mostrarEstatus('❌ Error: ' + error.message, 'error');
    }
}

/**
 * Inserta la fecha y hora en la celda activa
 */
async function insertarFechaHora() {
    console.log('Insertando fecha y hora...');
    
    try {
        await Excel.run(async (context) => {
            console.log('Excel.run iniciado (FechaHora)');
            
            const cell = context.application.activeCell;
            cell.load('address');
            
            await context.sync();
            console.log('Celda activa:', cell.address);
            
            const ahora = new Date();
            const valor = formatearFechaHora(ahora);
            
            console.log('Valor a insertar:', valor);
            
            cell.values = [[valor]];
            
            await context.sync();
            
            console.log('✅ Fecha y hora insertadas exitosamente');
            mostrarEstatus(`✅ Fecha y hora insertadas en ${cell.address}: ${valor}`, 'success');
        });
    } catch (error) {
        console.error('❌ Error al insertar fecha y hora:', error);
        mostrarEstatus('❌ Error: ' + error.message, 'error');
    }
}

/**
 * Funciones auxiliares de formato
 */

function formatearHora(fecha) {
    const horas = String(fecha.getHours()).padStart(2, '0');
    const minutos = String(fecha.getMinutes()).padStart(2, '0');
    const segundos = String(fecha.getSeconds()).padStart(2, '0');
    return `${horas}:${minutos}:${segundos}`;
}

function formatearHoraAmPm(fecha) {
    let horas = fecha.getHours();
    const minutos = String(fecha.getMinutes()).padStart(2, '0');
    const ampm = horas >= 12 ? 'PM' : 'AM';
    horas = horas % 12;
    horas = horas ? horas : 12;
    horas = String(horas).padStart(2, '0');
    return `${horas}:${minutos} ${ampm}`;
}

function formatearMinutosSegundos(fecha) {
    const minutos = String(fecha.getMinutes()).padStart(2, '0');
    const segundos = String(fecha.getSeconds()).padStart(2, '0');
    return `${minutos}:${segundos}`;
}

function formatearFechaHora(fecha) {
    const dia = String(fecha.getDate()).padStart(2, '0');
    const mes = String(fecha.getMonth() + 1).padStart(2, '0');
    const anio = fecha.getFullYear();
    const horas = String(fecha.getHours()).padStart(2, '0');
    const minutos = String(fecha.getMinutes()).padStart(2, '0');
    const segundos = String(fecha.getSeconds()).padStart(2, '0');
    return `${dia}/${mes}/${anio} ${horas}:${minutos}:${segundos}`;
}

/**
 * Muestra mensaje de estado
 */
function mostrarEstatus(mensaje, tipo = 'info') {
    console.log('Estado:', mensaje);
    
    const statusDiv = document.getElementById('status');
    const statusText = document.getElementById('statusText');
    
    if (!statusDiv || !statusText) {
        console.error('❌ Elementos de status no encontrados');
        return;
    }
    
    statusDiv.className = 'status-section ' + tipo;
    statusText.textContent = mensaje;
    
    // Auto-ocultar después de 5 segundos
    setTimeout(() => {
        statusText.textContent = '';
    }, 5000);
}

/**
 * Log para debugging
 */
console.log('✅ taskpane.js cargado correctamente');
console.log('Esperando Office.onReady...');
