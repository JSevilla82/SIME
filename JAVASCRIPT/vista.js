//Variables - Elementos del DOM

var urlSitio = `${_spPageContextInfo.webAbsoluteUrl}`;


function hideContainerInfo() {
 const containers = document.querySelectorAll('.containerInfo, .containerTablaUsuarios');
  containers.forEach(container => {
    container.style.display = 'none';
  });
}

function showContainerInfo() {
 const containers = document.querySelectorAll('.containerInfo, .containerTablaUsuarios');
  containers.forEach(container => {
    container.style.display = 'block';
  });
}


window.addEventListener('DOMContentLoaded', function () {

  hideContainerInfo()

});

function mostrarAlerta(titulo, texto, icono) {
  Swal.fire({
    title: titulo,
    text: texto,
    icon: icono,
    confirmButtonColor: '#0459A1',
    confirmButtonText: 'Aceptar'
  });
}

function mostrarMensajeExito(mensaje) {
  Swal.fire({
    title: 'Éxito',
    text: mensaje,
    icon: 'success',
    showConfirmButton: false,
    timer: 2000
  });
}

function mostrarTarjetaExito(mensaje) {
  Swal.fire({
    toast: true,
    position: 'top-end',  
    icon: 'success',
    title: mensaje,
    showConfirmButton: false,
    timer: 5000,  
    timerProgressBar: true,
    didOpen: (toast) => {
      toast.addEventListener('mouseenter', Swal.stopTimer);
      toast.addEventListener('mouseleave', Swal.resumeTimer);
    }
  });
}



function mostrarCargando(mensaje) {
  Swal.fire({
    html: `
      <div class="swal2-custom-loading">
        <div class="rotating-container">
          <img src="../../JairoSevilla/SiteAssets/InformesEntraID/IMAGENES/logo_carga.png" class="rotating-image" alt="Loading">
        </div>
        <div class="loading-message">${mensaje}</div>
      </div>
    `,
    allowOutsideClick: false,
    showConfirmButton: false,
    showCancelButton: false,
    customClass: {
      popup: 'swal2-noanimation',
      title: 'swal2-title-custom',
      htmlContainer: 'swal2-html-container-custom'
    }
  });
  setTimeout(() => document.body.focus(), 0);
}


function ocultarCargando() {
  Swal.close();
}


/* ---------------------------------- Tablas DataTables --------------------------------------- */


function mostrarLicenciasUsoTabla(datos) {
  if ($.fn.DataTable.isDataTable('#tablaEstadoLicencias')) {
    $('#tablaEstadoLicencias').DataTable().destroy();
  }

  $('#tablaEstadoLicencias tbody').empty();

  const columnas = [
    { data: 'licencia', title: 'Licencia', className: 'table-left-align' },
    { data: 'total', title: 'Total', className: 'table-left-align' },
    { data: 'uso', title: 'En Uso', className: 'table-left-align' },
    { data: 'disponibles', title: 'Disponibles', className: 'table-left-align' }
  ];

  $('#tablaEstadoLicencias').DataTable({
    data: datos,
    columns: columnas,
    paging: false,       // Desactiva el paginado
    searching: false,    // Desactiva el buscador
    info: false,         // Desactiva la información de registros
    ordering: true,      // Activa el ordenamiento
    order: [[1, 'desc']], // Ordena de mayor a menor según el total de licencias
    language: {
      url: '//cdn.datatables.net/plug-ins/1.10.21/i18n/Spanish.json'
    },
    dom: 't',    
    border: true,       
    rowCallback: function (row, data, index) {

      if (data.disponibles === 0) {
        $(row).css('background-color', '#FFCCCC'); // rojo suave
    } else if (data.disponibles < 2) {
        $(row).css('background-color', '#FFE5CC'); // naranja suave
    } else {
        $(row).css('background-color', '#CCFFCC'); // verde suave
    }
    
    }
  });
  generarGraficoLicencias(datos);
}

function actualizarEstadisticas(data) {
  const totalUsuarios = data.length;
  let usuariosConLicencia = 0;
  let usuariosSinLicencia = 0;
  let usuariosBloqueadosConLicencia = 0;

  data.forEach(usuario => {
      if (usuario.assignedLicenses === 'Sin Licencia') {
          usuariosSinLicencia++;
      } else {
          usuariosConLicencia++;

          if (usuario.accountEnabled === 'Bloqueado') {
              usuariosBloqueadosConLicencia++;
          }
      }
  });

  document.getElementById('totalUsuarios').textContent = totalUsuarios;
  document.getElementById('usuariosFiltrados').textContent = $('#tablaUsuarios').DataTable().rows({ filter: 'applied' }).count();
  document.getElementById('usuariosConLicencia').textContent = usuariosConLicencia;
  document.getElementById('usuariosSinLicencia').textContent = usuariosSinLicencia;
  document.getElementById('usuariosLicenciaBloqueado').textContent = usuariosBloqueadosConLicencia;

  if (document.getElementById('totalUsuarios').textContent === document.getElementById('usuariosFiltrados').textContent){
    document.getElementById('usuariosFiltrados').textContent = 0;
  }
}


function mostrarUsuariosTabla(datos) {
    if ($.fn.DataTable.isDataTable('#tablaUsuarios')) {
        $('#tablaUsuarios').DataTable().destroy();
    }

    $('#tablaUsuarios tbody').empty();

    const columnas = [
        { data: 'displayName', title: 'Nombre Completo' },
        { data: 'mail', title: 'Usuario/Correo Electrónico' },
        { data: 'jobTitle', title: 'Cargo' },
        { data: 'officeLocation', title: 'Ubicación' },
        {
            data: 'assignedLicenses',
            title: 'Licencias',
            render: function(data, type, row) {
                const licenciasPrincipales = ['Microsoft 365 E3', 'Office 365 E3', 'Office 365 F3'];
                const licencias = data.split(', ');
                const licenciasFormateadas = licencias.map(licencia => {
                    if (licenciasPrincipales.includes(licencia)) {
                        return `<strong>${licencia}</strong>`;
                    }
                    return licencia;
                });
                return licenciasFormateadas.join(', ');
            }
        },
        {
            data: 'accountEnabled',
            title: 'Inicio de Sesión',
            render: function(data, type, row) {
                if (data === 'Habilitado') {
                    return `<span style="color: green; font-weight: bold;">${data}</span>`;
                } else if (data === 'Bloqueado') {
                    return `<span style="color: red; font-weight: bold;">${data}</span>`;
                } else {
                    return data;
                }
            }
        },
        { data: 'createdDateTime', title: 'Creación' }
    ];

    $('#tablaUsuarios').DataTable({
        data: datos,
        columns: columnas,
        language: {
            url: '//cdn.datatables.net/plug-ins/1.10.21/i18n/Spanish.json'
        },
        rowCallback: function(row, data, index) {
            if (data.assignedLicenses !== 'Sin Licencia') {
                $(row).css('background-color', '#daf8d7');
            }
        },
        drawCallback: function(settings) {
            actualizarEstadisticas(datos);
        }
    });

    actualizarEstadisticas(datos);
}



function mostrarMovimientosLicenciaTabla(datos) {

  if ($.fn.DataTable.isDataTable('#tablaMovimientosLicencia')) {
    $('#tablaMovimientosLicencia').DataTable().destroy();
  }

  $('#tablaMovimientosLicencia tbody').empty();

  const columnas = [
    { data: 'id', title: 'ID' },
    { data: 'fechaHora', title: 'Fecha y Hora' },
    { data: 'usuarioAfectado', title: 'Usuario Afectado' },
    { data: 'actor', title: 'Actor' }
  ];

  $('#tablaMovimientosLicencia').DataTable({
    data: datos,
    columns: columnas,
    language: {
      url: '//cdn.datatables.net/plug-ins/1.10.21/i18n/Spanish.json'
    },
    responsive: true,
    columnDefs: [
      { targets: '_all', defaultContent: '' }
    ]
  });
}




