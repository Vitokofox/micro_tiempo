// Espera a que el DOM esté completamente cargado
document.addEventListener('DOMContentLoaded', function() {
    // Inicializar cuando Office esté listo
    Office.onReady((reason) => {
        if (reason === Office.HostType.Excel) {
            console.log('Excel Web Add-in cargado correctamente');
            inicializarComplemento();
        }
    });
});

function inicializarComplemento() {
    // Actualizar la hora cada segundo
    actualizarHoraDisplay();
    setInterval(actualizarHoraDisplay, 1000);
    
    // Asignar eventos a los botones
    document.getElementById('btnInsertarHora').onclick = insertarHora;
    document.getElementById('btnInsertarFecha').onclick = insertarFechaHora;
}

/**
 * Actualiza la pantalla de la hora actual
 */
function actualizarHoraDisplay() {
    const ahora = new Date();
    const horas = String(ahora.getHours()).padStart(2, '0');
    const minutos = String(ahora.getMinutes()).padStart(2, '0');
    const segundos = String(ahora.getSeconds()).padStart(2, '0');
    
    const horaFormato = `${horas}:${minutos}:${segundos}`;
    document.getElementById('horaActual').textContent = horaFormato;
}

/**
 * Inserta la hora en la celda activa (según formato seleccionado)
 */
async function insertarHora() {
    try {
        const formato = document.querySelector('input[name="formato"]:checked').value;
        
        await Excel.run(async (context) => {
            const cell = context.application.activeCell;
            cell.load('address');
            
            // Obtener la hora actual
            const ahora = new Date();
            let valor;
            
            switch(formato) {
                case 'hora':
                    // hh:mm:ss
                    valor = formatearHora(ahora);
                    break;
                case 'horaAmPm':
                    // hh:mm AM/PM
                    valor = formatearHoraAmPm(ahora);
                    break;
                case 'horaMinutos':
                    // mm:ss
                    valor = formatearMinutosSegundos(ahora);
                    break;
                case 'fechaHora':
                    // dd/mm/yyyy hh:mm:ss
                    valor = formatearFechaHora(ahora);
                    break;
                default:
                    valor = formatearHora(ahora);
            }
            
            // Insertar valor en la celda activa
            cell.values = [[valor]];
            
            await context.sync();
            
            mostrarEstatus(`✅ Hora insertada en celda ${cell.address}`, 'success');
            console.log('Hora insertada:', valor);
        });
    } catch (error) {
        console.error('Error al insertar hora:', error);
        mostrarEstatus('❌ Error: ' + error.message, 'error');
    }
}

/**
 * Inserta la fecha y hora en la celda activa
 */
async function insertarFechaHora() {
    try {
        await Excel.run(async (context) => {
            const cell = context.application.activeCell;
            cell.load('address');
            
            const ahora = new Date();
            const valor = formatearFechaHora(ahora);
            
            cell.values = [[valor]];
            
            await context.sync();
            
            mostrarEstatus(`✅ Fecha y hora insertadas en ${cell.address}`, 'success');
        });
    } catch (error) {
        console.error('Error:', error);
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
    const statusDiv = document.getElementById('status');
    const statusText = document.getElementById('statusText');
    
    statusDiv.className = 'status-section ' + tipo;
    statusText.textContent = mensaje;
    
    // Auto-ocultar después de 5 segundos
    setTimeout(() => {
        statusText.textContent = '';
    }, 5000);
}
