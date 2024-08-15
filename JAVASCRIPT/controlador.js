
function generarGraficoLicencias(datos) {

  console.log(datos);
  datos.sort((a, b) => b.disponibles - a.disponibles);

  const ctx = document.getElementById('graficoLicencias').getContext('2d');
  const labels = datos.map(item => item.licencia);
  const disponibles = datos.map(item => item.disponibles);

  const chartColors = {
    disponibles: 'rgba(4, 89, 161, 0.6)'
  };

  new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        {
          label: 'Licencias Disponibles',
          data: disponibles,
          backgroundColor: chartColors.disponibles
        }
      ]
    },
    options: {
      responsive: true,
      scales: {
        x: {
          stacked: false
        },
        y: {
          stacked: false,
          beginAtZero: true,
          suggestedMax: Math.max(...disponibles) + 2
        }
      },
      plugins: {
        legend: {
          position: 'top'
        },
        title: {
          display: false,
          text: 'Licencias Disponibles'
        },
        tooltip: {
          callbacks: {
            label: function (context) {
              return context.dataset.label + ': ' + context.raw;
            }
          }
        }
      }
    },
    plugins: [{
      id: 'customPlugin',
      afterDatasetsDraw: function (chart) {
        const ctx = chart.ctx;
        chart.data.datasets.forEach(function (dataset, i) {
          const meta = chart.getDatasetMeta(i);
          meta.data.forEach(function (bar, index) {
            const data = dataset.data[index];
            ctx.fillStyle = 'black';
            ctx.fillText(data, bar.x, bar.y - 5);
          });
        });
      }
    }]
  });
}




function contarLicencias(datosUsuarios) {

  const licenciasSet = new Set([

    'Microsoft 365 E3',
    'Office 365 E3',
    'Office 365 F3',
    'Power BI Pro',
    'Power BI Premium por Usuario',
    'Power Apps Premium',
    'Project Plan 3',
    'Salas de Microsoft Teams Pro',
    'Power Automate per user plan'
  ]);

  const resultado = Array.from(licenciasSet).reduce((acc, licencia) => {
    acc[licencia] = 0;
    return acc;
  }, {});

  datosUsuarios.forEach(usuario => {
    const licencias = usuario.assignedLicenses.split(', ');
    licencias.forEach(licencia => {
      if (licenciasSet.has(licencia)) {
        resultado[licencia]++;
      }
    });
  });

  return resultado;
}



function formatearFechaColombia(fechaISO) {
  const opciones = {
    timeZone: 'America/Bogota',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false
  };

  const fecha = new Date(fechaISO);
  const formateada = new Intl.DateTimeFormat('es-CO', opciones).format(fecha);

  return formateada;
}

function prepararDatosParaTabla(datosUsuarios, licenciasTotales) {

  const resultadoLicenciasUsuarios = contarLicencias(datosUsuarios);

  const datosTabla = Object.keys(resultadoLicenciasUsuarios).map(licencia => {
    const licenciaTotal = licenciasTotales.find(l => l.nombre === licencia);
    const licenciaEnUso = resultadoLicenciasUsuarios[licencia];

    if (licenciaTotal) {
      const disponibles = licenciaEnUso > 0 ? licenciaTotal.total - licenciaEnUso : licenciaTotal.total;

      return {
        licencia,
        disponibles,
        uso: licenciaEnUso,
        total: licenciaTotal.total
      };
    } else {
      return {
        licencia,
        disponibles: 0,
        uso: licenciaEnUso,
        total: 0
      };
    }
  });

  return datosTabla;
}


let datosOriginales = [];

function aplicarFiltro() {
  const filtroSeleccionado = document.querySelector('input[name="opcion"]:checked').value;
  let datosFiltrados = datosOriginales;

  if (filtroSeleccionado === 'Con licencia') {
    datosFiltrados = datosOriginales.filter(usuario => usuario.assignedLicenses && usuario.assignedLicenses.trim() !== 'Sin Licencia');
  } else if (filtroSeleccionado === 'Sin licencia') {
    datosFiltrados = datosOriginales.filter(usuario => !usuario.assignedLicenses || usuario.assignedLicenses.trim() === 'Sin Licencia');
  } else if (filtroSeleccionado !== 'Todos') {
    datosFiltrados = datosOriginales.filter(usuario => usuario.assignedLicenses.includes(filtroSeleccionado));
  }

  mostrarUsuariosTabla(datosFiltrados);
}


function generarOpcionesFiltroLicencias(usuarios) {
  const filtroLicencias = document.getElementById('filtroLicencias');
  filtroLicencias.innerHTML = '';

  const defaultOption = document.createElement('option');
  defaultOption.value = 'seleccionar';
  defaultOption.textContent = 'Todas las licencias';
  filtroLicencias.appendChild(defaultOption);

  const licenciasPago = document.createElement('optgroup');
  licenciasPago.label = 'Licencias de Pago';

  const licenciasFree = document.createElement('optgroup');
  licenciasFree.label = 'Licencias Free/Prueba';

  const licenciasPagoEspecificas = new Set([
    'Microsoft 365 E3',
    'Office 365 E3',
    'Office 365 F3',
    'Project Plan 3',
    'Power BI Pro',
    'Power BI Premium por Usuario',
    'Salas de Microsoft Teams Pro',
    'Power Automate per user plan',
    'Power Apps Premium'
  ]);

  const licenciasSet = new Set();

  usuarios.forEach(usuario => {
    usuario.assignedLicenses.split(',').map(licencia => licencia.trim()).forEach(licencia => {
      if (licencia && licencia !== 'Sin Licencia') {
        licenciasSet.add(licencia);
      }
    });
  });

  Array.from(licenciasSet).sort().forEach(licencia => {
    const option = document.createElement('option');
    option.value = licencia;
    option.textContent = licencia;

    if (licenciasPagoEspecificas.has(licencia)) {
      licenciasPago.appendChild(option);
    } else {
      licenciasFree.appendChild(option);
    }
  });

  filtroLicencias.appendChild(licenciasPago);
  filtroLicencias.appendChild(licenciasFree);
}



document.getElementById('filtroLicencias').addEventListener('change', function () {
  aplicarFiltroCombinado();
});

document.getElementById('filtroUsuarios').addEventListener('change', function () {
  aplicarFiltroCombinado();
});

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function aplicarFiltroCombinado() {
  const licenciaSeleccionada = document.getElementById('filtroLicencias').value;
  const filtroUsuarios = document.getElementById('filtroUsuarios').value;
  const dataTable = $('#tablaUsuarios').DataTable();
  const columnaLicencias = 4;


  dataTable.search('').column(columnaLicencias).search('').draw();


  if (filtroUsuarios === 'todos') {
    if (licenciaSeleccionada === 'seleccionar') {
      dataTable.search('').column(columnaLicencias).search('').draw();
    } else {
      const escapedLicencia = escapeRegExp(licenciaSeleccionada);
      const regex = new RegExp(`(^|,)\\s*${escapedLicencia}\\s*(,|$)`, 'i');
      dataTable.column(columnaLicencias).search(regex.source, true, false).draw();
    }
  }

  else if (filtroUsuarios === 'conLicencia') {
    if (licenciaSeleccionada === 'seleccionar') {
      dataTable.column(columnaLicencias).search('^(?!Sin Licencia$).*$', true, false).draw(); // Mostrar solo usuarios con licencias
    } else {
      const escapedLicencia = escapeRegExp(licenciaSeleccionada);
      const regex = new RegExp(`(^|,)\\s*${escapedLicencia}\\s*(,|$)`, 'i');
      dataTable.column(columnaLicencias).search(regex.source, true, false).draw(); // Filtrar por licencia seleccionada
    }
  }

  else if (filtroUsuarios === 'sinLicencia') {
    dataTable.column(columnaLicencias).search('^Sin Licencia$', true, false).draw(); // Mostrar solo usuarios sin licencia
  }
}



(function () {
  const filtroUsuarios = document.getElementById('filtroUsuarios');
  const filtroLicencias = document.getElementById('filtroLicencias');
  const opcionSinLicencia = filtroUsuarios.querySelector('option[value="sinLicencia"]');

  filtroUsuarios.addEventListener('change', function () {
    if (filtroUsuarios.value === 'sinLicencia') {
      if (filtroLicencias.value !== 'seleccionar') {
        filtroLicencias.value = 'seleccionar';
      }
      filtroLicencias.disabled = true;
    } else {
      filtroLicencias.disabled = false;
    }
  });

  filtroLicencias.addEventListener('change', function () {
    if (filtroLicencias.value !== 'seleccionar') {
      opcionSinLicencia.disabled = true;

      if (filtroUsuarios.value === 'sinLicencia') {
        filtroUsuarios.value = 'todos';
      }
    } else {
      opcionSinLicencia.disabled = false;
    }
  });
})();



window.addEventListener('DOMContentLoaded', async function () {
  try {
    $('#tablaDatos').DataTable().destroy();

    mostrarCargando('Descargando datos, por favor espere...');

    await inicializarAutenticacion();
    const token = await obtenerToken();

    const datosUsuarios = await obtenerUsuarios(token);
    const licenciasTotales = await obtenerLicenciasDisponibles(token);

    const datosTabla = prepararDatosParaTabla(datosUsuarios, licenciasTotales);

    showContainerInfo();
    mostrarLicenciasUsoTabla(datosTabla);

    generarOpcionesFiltroLicencias(datosUsuarios);
    mostrarUsuariosTabla(datosUsuarios);

    const movimientosLicencias = await fetchAuditLogs(token);
    const movimientosLicenciasDepurado = extraerInfoMovimientosLicencias(movimientosLicencias);
    mostrarMovimientosLicenciaTabla(movimientosLicenciasDepurado);

    ocultarCargando();
    mostrarTarjetaExito("La transferencia de datos se ha completado exitosamente.");

  } catch (error) {
    console.error('Error general:', error);
    ocultarCargando();
  }
});




document.getElementById('exportBtn').addEventListener('click', function (event) {
  event.preventDefault();

  var table = $('#tablaUsuarios').DataTable();
  var data = table.rows().data().toArray();

  mostrarTarjetaExito("Los registros se han exportado correctamente.");

  setTimeout(() => {
    exportarExcelDeTabla(data, 'Usuarios.xlsx');
  }, 1000);

});

document.getElementById('sincronizarBtn').addEventListener('click', async function (event) {

  event.preventDefault();

  const icono = document.getElementById('iconoSincronizar');
  icono.classList.add('rotar');

  try {
    const token = await obtenerToken();
    const datosUsuarios = await obtenerUsuarios(token);
    generarOpcionesFiltroLicencias(datosUsuarios);
    mostrarUsuariosTabla(datosUsuarios);

    mostrarTarjetaExito("Sincronización de perfiles realizada correctamente.");

  } catch (error) {
    console.error('Error en la sincronización:', error);
    mostrarAlerta('Error', 'Hubo un problema en la sincronización. Intenta nuevamente más tarde.', 'error');
  } finally {
    icono.classList.remove('rotar');  // Remover clase de rotación
  }
});




function procesarLicenciasPrincipales(datosUsuarios) {
  return datosUsuarios.map(usuario => {
    const licenciasPrincipales = [];
    const otrasLicencias = [];

    usuario.assignedLicenses.split(', ').forEach(licencia => {
      if (['Microsoft 365 E3', 'Office 365 E3', 'Office 365 F3'].includes(licencia)) {
        licenciasPrincipales.push(licencia);
      } else {
        otrasLicencias.push(licencia);
      }
    });

    return {
      ...usuario,
      assignedLicenses: otrasLicencias.length > 0 ? otrasLicencias.join(', ') : 'Sin Licencia',
      licenciaPrincipal: licenciasPrincipales.length > 0 ? licenciasPrincipales.join(', ') : 'Sin Licencia'
    };
  });
}

function exportarExcelDeTabla(datosTabla) {

  const datosProcesados = procesarLicenciasPrincipales(datosTabla);
  var wb = XLSX.utils.book_new();

  var datosFormateados = datosProcesados.map(function (row) {
    return {
      'NOMBRE COMPLETO': row.displayName === '-' ? '' : row.displayName,
      'USUARIO': row.mail === '-' ? '' : row.mail,
      'CARGO': row.jobTitle === '-' ? '' : row.jobTitle,
      'UBICACIÓN': row.officeLocation === '-' ? '' : row.officeLocation,
      'LICENCIA PRINCIPAL': row.licenciaPrincipal === '-' ? '' : row.licenciaPrincipal,
      'OTRAS LICENCIAS': row.assignedLicenses === '-' ? '' : row.assignedLicenses,
      'INICIO DE SESIÓN': row.accountEnabled === '-' ? '' : row.accountEnabled,
      'CREACIÓN': row.createdDateTime === '-' ? '' : row.createdDateTime
    };
  });

  var ws = XLSX.utils.json_to_sheet(datosFormateados);

  var colWidths = [
    { wpx: 200 }, // Nombre Completo
    { wpx: 180 }, // Usuario
    { wpx: 150 }, // Cargo
    { wpx: 150 }, // Ubicación
    { wpx: 200 }, // Licencia Principal
    { wpx: 200 }, // Otras Licencias
    { wpx: 120 }, // Inicio de Sesión
    { wpx: 150 }  // Creación
  ];
  ws['!cols'] = colWidths;

  XLSX.utils.book_append_sheet(wb, ws, 'Usuarios');

  var now = new Date();
  var year = now.getFullYear();
  var month = (now.getMonth() + 1).toString().padStart(2, '0');
  var day = now.getDate().toString().padStart(2, '0');
  var hours = now.getHours().toString().padStart(2, '0');
  var minutes = now.getMinutes().toString().padStart(2, '0');
  var seconds = now.getSeconds().toString().padStart(2, '0');
  var formattedDate = `${year}${month}${day}_${hours}${minutes}${seconds}`;
  var nombreArchivo = 'Usuarios_SIME_' + formattedDate + '.xlsx';

  XLSX.writeFile(wb, nombreArchivo);

}






