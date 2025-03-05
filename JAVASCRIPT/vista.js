//Variables - Elementos del DOM

const contenedor = document.getElementById('resultadoRevisionArchivo');
const animacion = document.getElementById('animacionResultado');
const mensaje = document.getElementById('mensajeResultado');

const filtroUsuarios = document.getElementById('filtroUsuarios');
const filtroLicencias = document.getElementById('filtroLicencias');
const opcionSinLicencia = filtroUsuarios.querySelector('option[value="sinLicencia"]');

const nombreArchivo2FA = document.getElementById("nombreArchivo2FA");
const inputArchivo2FA = document.getElementById("archivoUsuario2FA");
const botonCargarArchivo2FA = document.querySelector(".cargarArchivo");
const botonIniciarActualizacion2FA = document.getElementById("btnIniciarActualizacion2FA");

var urlSIME = `${_spPageContextInfo.webAbsoluteUrl}`;
document.addEventListener('contextmenu', event => event.preventDefault());


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


document.addEventListener('DOMContentLoaded', function () {
  const titulo = document.querySelector('.titulo h1');
  if (titulo) {
    titulo.style.marginLeft = '-100px';
  }
});


window.addEventListener('DOMContentLoaded', function () {

  hideContainerInfo();
  eventoMantenerLogin();
  controlarBotonesExcluir2FA();

  document.querySelectorAll('.tab-button').forEach(button => {
    button.addEventListener('click', (event) => {

      event.preventDefault();

      document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
      document.querySelectorAll('.tab-content-item').forEach(item => item.classList.remove('active'));

      button.classList.add('active');
      document.getElementById(button.dataset.tab).classList.add('active');
    });
  });

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


function mostrarResultado(isSuccess) {
  // Mostrar el contenedor
  contenedor.style.display = 'flex';

  if (isSuccess) {
    animacion.textContent = '✓';
    animacion.className = 'animacion ok visible';
    mensaje.textContent = 'El archivo cumple con todos los requisitos.';
    mensaje.className = 'mensaje visible';
  } else {
    animacion.textContent = '✗';
    animacion.className = 'animacion error visible';
    mensaje.textContent = 'El archivo no es compatible y no se puede procesar.';
    mensaje.className = 'mensaje visible';
  }
}

function ocultarResultado() {
  const contenedor = document.getElementById('resultadoRevisionArchivo');
  const animacion = document.getElementById('animacionResultado');
  const mensaje = document.getElementById('mensajeResultado');

  // Ocultar los elementos
  if (animacion && mensaje && contenedor) {
    animacion.className = 'animacion';
    mensaje.className = 'mensaje';

    // Ocultar el contenedor
    contenedor.style.display = 'none';
  } else {
    console.error("No se encontraron los elementos para ocultar el resultado.");
  }
}

/* ---------------------------------- Tablas DataTables --------------------------------------- */

function mostrarLicenciasUsoTabla(datos) {

  if ($.fn.DataTable.isDataTable('#tablaEstadoLicencias')) {
    $('#tablaEstadoLicencias').DataTable().destroy();
  }

  $('#tablaEstadoLicencias tbody').empty();

  const totalLicenciasExpiradas = datos.reduce((sum, item) => sum + item.licenciasExpiradas, 0);
  const hayLicenciasExpiradas = datos.some(item => item.licenciasExpiradas > 0);

  const columnas = [
    {
      data: 'licencia',
      title: 'Licencia',
      className: 'table-left-align'
    },
    {
      data: 'total',
      title: 'Compradas',
      className: 'table-left-align',
      render: (data, type, row) => data < 0 ? 0 : data
    },
    ...(hayLicenciasExpiradas ? [{
      data: 'licenciasExpiradas',
      title: 'Expiradas',
      className: 'table-left-align',
      render: (data, type, row) => data < 0 ? 0 : data
    }] : []),
    ...(hayLicenciasExpiradas ? [{
      data: null,
      title: 'Total Licencias',
      className: 'table-left-align',
      render: (data, type, row) => {
        const totalLic = row.total + row.licenciasExpiradas;
        return totalLic < 0 ? 0 : totalLic;
      },
      visible: hayLicenciasExpiradas // Esta columna solo se muestra si aplica
    }] : []),
    {
      data: 'uso',
      title: 'Asignadas',
      className: 'table-left-align',
      render: (data, type, row) => data < 0 ? 0 : data
    },
    {
      data: 'disponibles',
      title: 'Disponibles',
      className: 'table-left-align',
      render: (data, type, row) => data < 0 ? 0 : data
    }
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
      if (data.licenciasExpiradas > 0) {
        $(row).css('background-color', '#FF9999'); // rojo más intenso para licencias expiradas
        mostrarAlerta('Licencias expiradas', `Hay ${totalLicenciasExpiradas} licencias expiradas, póngase en contacto con su partner.`, 'warning');
      } else if (data.total === 0) {
        $(row).css('background-color', '#FF9999'); // rojo más intenso para asignadas sin compradas
      } else if (data.disponibles === 0) {
        $(row).css('background-color', '#FFCCCC'); // rojo suave
      } else if (data.disponibles < 2) {
        $(row).css('background-color', '#FFE5CC'); // naranja suave
      } else {
        $(row).css('background-color', '#CCFFCC'); // verde suave
      }
    }

  });

  if (totalLicenciasExpiradas > 0) {
    $('#tablaEstadoLicencias').append(`
      <tfoot>
        <tr>
          <td colspan="${columnas.length}" style="font-weight: bold; color: red;">
            Hay <strong>${totalLicenciasExpiradas}</strong> licencias expiradas, póngase en contacto con su partner.
          </td>
        </tr>
      </tfoot>
    `);
  }

  generarGraficoLicencias(datos);
}


function actualizarEstadisticas(data) {
  const totalUsuarios = data.length;
  let usuariosConLicencia = 0;
  let usuariosBloqueadosConLicencia = 0;

  data.forEach(usuario => {
    if (usuario.assignedLicenses !== 'Sin Licencia') {
      usuariosConLicencia++;

      if (usuario.accountEnabled === 'Bloqueado') {
        usuariosBloqueadosConLicencia++;
      }
    }
  });

  document.getElementById('totalUsuarios').textContent = totalUsuarios;
  document.getElementById('usuariosFiltrados').textContent = $('#tablaUsuarios').DataTable().rows({ filter: 'applied' }).count();
  document.getElementById('usuariosConLicencia').textContent = usuariosConLicencia;
  document.getElementById('usuariosLicenciaBloqueado').textContent = usuariosBloqueadosConLicencia;

  if (document.getElementById('totalUsuarios').textContent === document.getElementById('usuariosFiltrados').textContent) {
    document.getElementById('usuariosFiltrados').textContent = 0;
  }
}

/*
function mostrarUsuariosTabla(datos) {
  if ($.fn.DataTable.isDataTable('#tablaUsuarios')) {
    $('#tablaUsuarios').DataTable().destroy();
  }

  $('#tablaUsuarios tbody').empty();

  const columnas = [
    { data: 'displayName', title: 'Nombre' },
    { data: 'mail', title: 'Correo Electrónico' },
    { data: 'jobTitle', title: 'Cargo' },
    { data: 'officeLocation', title: 'Oficina' },

    {
      data: 'assignedLicenses',
      title: 'Licencias Asignadas',
      render: function (data, type, row) {

        const licenciasPrincipales = Object.values(window.BDIdentificadorLicencias)
          .filter(licencia => licencia.LicenciaPrincipal === 1)
          .map(licencia => licencia.NombreLicencia);

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
      render: function (data, type, row) {
        if (data === 'Habilitado') {
          return `<span style="color: green; font-weight: bold;">${data}</span>`;
        } else if (data === 'Bloqueado') {
          return `<span style="color: red; font-weight: bold;">${data}</span>`;
        } else {
          return data;
        }
      }
    },
    { data: 'createdDateTime', title: 'Creación del usuario' },
    // {
    //   data: 'lastSignInDateTime',
    //   title: 'Ultima conexión'
    // }

    {
      data: 'lastSignInDateTime',
      title: 'Última conexión',
      render: function (data, type, row) {
        
        const datePattern = /^\d{2}\/\d{2}\/\d{4}, \d{2}:\d{2}:\d{2}$/;
        if (type === 'display' && datePattern.test(data)) {
        
          const [datePart, timePart] = data.split(', ');
          const [day, month, year] = datePart.split('/');
          const formattedDate = new Date(`${year}-${month}-${day}T${timePart}`);

          const currentDate = new Date();
          const diffInDays = (currentDate - formattedDate) / (1000 * 60 * 60 * 24);

          let color;
          if (diffInDays > 30) {
            color = '#B22222'; // Rojo oscuro
          } else if (diffInDays > 7) {
            color = '#FF8C00'; // Naranja oscuro
          } else {
            color = '#006400'; // Verde oscuro
          }
          

          return `<span style="color:${color}; font-weight: bold;">${data}</span>`;
        }
        return data || '';
      }
    }

  ];

  $('#tablaUsuarios').DataTable({
    data: datos,
    columns: columnas,
    language: {
      url: '//cdn.datatables.net/plug-ins/1.10.21/i18n/Spanish.json'
    },
    rowCallback: function (row, data, index) {
      if (data.assignedLicenses !== 'Sin Licencia') {
        $(row).css('background-color', '#daf8d7');
      }else{
        $(row).css('background-color', '#f8d7da');
      }
    },
    drawCallback: function (settings) {
      actualizarEstadisticas(datos);
    }
  });

  actualizarEstadisticas(datos);
}
*/

function mostrarUsuariosTabla(datos) {
  if ($.fn.DataTable.isDataTable('#tablaUsuarios')) {
    $('#tablaUsuarios').DataTable().destroy();
  }
  $('#tablaUsuarios tbody').empty();

  // Función para eliminar acentos
  function quitarAcentos(texto) {
    return texto
      .normalize('NFD') // Descompone los caracteres con acentos en sus componentes base
      .replace(/[\u0300-\u036f]/g, '') // Elimina los diacríticos (acentos)
      .toLowerCase(); // Convierte a minúsculas
  }

  // Función para detectar si un texto parece ser un correo electrónico
  function esCorreoElectronico(texto) {
    const correoRegex = /^[^@]+@[^@]+\.[^@]+$/;
    return correoRegex.test(texto);
  }

  // Función para extraer el usuario de un correo electrónico
  function obtenerUsuarioCorreo(correo) {
    if (correo && correo.includes('@')) {
      return correo.split('@')[0].toLowerCase(); // Extrae la parte antes del '@'
    }
    return correo ? correo.toLowerCase() : ''; // Si no es un correo, devuelve el texto completo
  }

  const columnas = [
    { data: 'displayName', title: 'Nombre' },
    { data: 'mail', title: 'Correo Electrónico' },
    { data: 'jobTitle', title: 'Cargo' },
    { data: 'officeLocation', title: 'Oficina' },
    {
      data: 'assignedLicenses',
      title: 'Licencias Asignadas',
      render: function (data, type, row) {
        const licenciasPrincipales = Object.values(window.BDIdentificadorLicencias)
          .filter(licencia => licencia.LicenciaPrincipal === 1)
          .map(licencia => licencia.NombreLicencia);
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
      render: function (data, type, row) {
        if (data === 'Habilitado') {
          return `<span style="color: green; font-weight: bold;">${data}</span>`;
        } else if (data === 'Bloqueado') {
          return `<span style="color: red; font-weight: bold;">${data}</span>`;
        } else {
          return data;
        }
      }
    },
    { data: 'createdDateTime', title: 'Creación del usuario' },
    {
      data: 'lastSignInDateTime',
      title: 'Última conexión',
      render: function (data, type, row) {
        const datePattern = /^\d{2}\/\d{2}\/\d{4}, \d{2}:\d{2}:\d{2}$/;
        if (type === 'display' && datePattern.test(data)) {
          const [datePart, timePart] = data.split(', ');
          const [day, month, year] = datePart.split('/');
          const formattedDate = new Date(`${year}-${month}-${day}T${timePart}`);
          const currentDate = new Date();
          const diffInDays = (currentDate - formattedDate) / (1000 * 60 * 60 * 24);
          let color;
          if (diffInDays > 30) {
            color = '#B22222'; // Rojo oscuro
          } else if (diffInDays > 7) {
            color = '#FF8C00'; // Naranja oscuro
          } else {
            color = '#006400'; // Verde oscuro
          }
          return `<span style="color:${color}; font-weight: bold;">${data}</span>`;
        }
        return data || '';
      }
    }
  ];

  $('#tablaUsuarios').DataTable({
    data: datos,
    columns: columnas,
    language: {
      url: '//cdn.datatables.net/plug-ins/1.10.21/i18n/Spanish.json'
    },
    rowCallback: function (row, data, index) {
      if (data.assignedLicenses !== 'Sin Licencia') {
        $(row).css('background-color', '#daf8d7');
      } else {
        $(row).css('background-color', '#f8d7da');
      }
    },
    drawCallback: function (settings) {
      actualizarEstadisticas(datos);
    },
    initComplete: function () {
      const api = this.api();

      // Personaliza el motor de búsqueda global
      $('#tablaUsuarios_filter input').off().on('keyup change clear', function () {
        const searchTerm = $(this).val(); // Obtiene el término de búsqueda
        const normalizedSearchTerm = quitarAcentos(searchTerm);

        // Limpia todos los filtros antes de aplicar uno nuevo
        api.columns().search('').draw();

        // Verifica si el término de búsqueda parece ser un correo electrónico
        if (normalizedSearchTerm.trim() !== '') {
          if (esCorreoElectronico(normalizedSearchTerm)) {
            // Aplica la lógica especial para correos electrónicos
            const searchUserPart = obtenerUsuarioCorreo(normalizedSearchTerm);
            api.column(1).search(searchUserPart, true, false).draw(); // Busca solo en la columna de correo
          } else {
            // Realiza una búsqueda normal en todas las columnas
            api.search(normalizedSearchTerm).draw();
          }
        } else {
          // Si el campo de búsqueda está vacío, muestra todos los registros
          api.search('').draw();
        }
      });
    },
    search: {
      regex: false, // Desactiva la búsqueda por expresiones regulares
    },
    createdRow: function (row, data, dataIndex) {
      // Normaliza los datos al crear cada fila
      $(row).data('normalized-data', quitarAcentos(JSON.stringify(data)));
    }
  });

  actualizarEstadisticas(datos);
}



function actualizarEstadisticas2FA(datosTotales, datos) {

  // Función para actualizar contadores
  function contarUsuariosPorEstado(datos) {
    let total = 0;
    let habilitado = 0;
    let deshabilitado = 0;
    let desconocido = 0;
    let excluido = 0;

    datos.forEach(usuario => {
      if (usuario.ExcluirCuenta === "1") {
        // Contar como excluido si el estado de exclusión es "1"
        excluido++;
      } else {
        // Si no está excluido, contar el estado 2FA
        if (usuario.estado2FA === 'Habilitado') {
          habilitado++;
        } else if (usuario.estado2FA === 'Deshabilitado') {
          deshabilitado++;
        } else if (usuario.estado2FA === 'Desconocido') {
          desconocido++;
        }
      }
    });

    // Sumar excluidos al total de usuarios
    total = habilitado + deshabilitado + desconocido + excluido;

    return { total, habilitado, deshabilitado, desconocido, excluido };
  }

  // Contar para todos los usuarios
  const totales = contarUsuariosPorEstado(datosTotales);
  document.getElementById('totalUsuarios2FA').textContent = totales.total;
  document.getElementById('habilitadoUsuarios2FA').textContent = totales.habilitado;
  document.getElementById('inhabilitadoUsuarios2FA').textContent = totales.deshabilitado;
  document.getElementById('DesconocidoUsuarios2FA').textContent = totales.desconocido;
  document.getElementById('excluidasUsuarios2FA').textContent = totales.excluido;

  // Contar para usuarios de la organización icontecOrg
  const org = contarUsuariosPorEstado(datos.icontecOrg);
  document.getElementById('totalUsuarios2FA_ORG').textContent = org.total;
  document.getElementById('habilitadoUsuarios2FA_ORG').textContent = org.habilitado;
  document.getElementById('deshabilitadoUsuarios2FA_ORG').textContent = org.deshabilitado;
  document.getElementById('desconocidoUsuarios2FA_ORG').textContent = org.desconocido;

  // Contar para usuarios de la organización icontecNet
  const net = contarUsuariosPorEstado(datos.icontecNet);
  document.getElementById('totalUsuarios2FA_NET').textContent = net.total;
  document.getElementById('habilitadoUsuarios2FA_NET').textContent = net.habilitado;
  document.getElementById('deshabilitadoUsuarios2FA_NET').textContent = net.deshabilitado;
  document.getElementById('desconocidoUsuarios2FA_NET').textContent = net.desconocido;

  // Contar para usuarios de la organización otros
  const otros = contarUsuariosPorEstado(datos.otros);
  document.getElementById('totalUsuarios2FA_OTROS').textContent = otros.total;
  document.getElementById('habilitadoUsuarios2FA_OTROS').textContent = otros.habilitado;
  document.getElementById('deshabilitadoUsuarios2FA_OTROS').textContent = otros.deshabilitado;
  document.getElementById('desconocidoUsuarios2FA_OTROS').textContent = otros.desconocido;
}

function mostrarUsuarios2FATabla(datos) {

  function inicializarDataTable(tablaId, data) {
    if ($.fn.DataTable.isDataTable(tablaId)) {
      $(tablaId).DataTable().destroy();
    }

    $(tablaId + ' tbody').empty();

    const columnas = [
      { data: 'displayName', title: 'Nombre Completo' },
      { data: 'mail', title: 'Usuario/Correo Electrónico' },
      { data: 'jobTitle', title: 'Cargo' },
      {
        data: 'accountEnabled',
        title: 'Inicio de Sesión',
        render: function (data) {
          if (data === 'Habilitado') {
            return `<span style="color: green; font-weight: bold;">${data}</span>`;
          } else if (data === 'Bloqueado') {
            return `<span style="color: red; font-weight: bold;">${data}</span>`;
          } else {
            return data;
          }
        }
      },
      {
        data: 'estado2FA',
        title: 'Estado 2FA',
        render: function (data) {
          if (data === 'Habilitado') {
            return `<span style="color: green; font-weight: bold;">${data}</span>`;
          } else if (data === 'Deshabilitado') {
            return `<span style="color: red; font-weight: bold;">${data}</span>`;
          } else if (data === 'Desconocido') {
            return `<span style="color: gray; font-weight: bold;">${data}</span>`;
          } else {
            return data;
          }
        }
      }
    ];

    $(tablaId).DataTable({
      data: data,
      columns: columnas,
      language: {
        url: '//cdn.datatables.net/plug-ins/1.10.21/i18n/Spanish.json'
      },
      order: [[4, 'asc']], // Sort by the "Estado 2FA" column (index 4) in ascending order
      rowCallback: function (row, data) {
        if (data.estado2FA === 'Habilitado') {
          $(row).css('background-color', '#e0ffe0');
        } else if (data.estado2FA === 'Deshabilitado') {
          $(row).css('background-color', '#ffe0e0');
        }
      }
    });
  }

  // Inicializar los DataTables
  inicializarDataTable('#2FAIcontecOrg', datos.icontecOrg);
  inicializarDataTable('#2FAIcontecNet', datos.icontecNet);
  inicializarDataTable('#2FAIcontecOtros', datos.otros);
  inicializarDataTable('#2FAIcontecExcluidos', datos.excluidos);
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

// Función para deshabilitar todos los botones
function deshabilitarBotones() {
  document.getElementById('sincronizarFechaSesionBtn').disabled = true;
  document.getElementById('sincronizarBtn').disabled = true;
  document.getElementById('exportBtn').disabled = true;
}

// Función para habilitar todos los botones
function habilitarBotones() {
  document.getElementById('sincronizarFechaSesionBtn').disabled = false;
  document.getElementById('sincronizarBtn').disabled = false;
  document.getElementById('exportBtn').disabled = false;
}

function reiniciarFiltrosUsuarios() {
  filtroUsuarios.value = 'todos';
  filtroLicencias.disabled = false;
  const opcionSinLicencia = filtroUsuarios.querySelector('option[value="sinLicencia"]');
  opcionSinLicencia.disabled = false;
}

const btnExcluir = document.getElementById('btnExcluir');
const btnReincluir = document.getElementById('btnReincluir');
const usuarioSelect = document.getElementById('usuarioSelect');
const btnReiniciar2FA = document.getElementById('btnResetear2FA'); //PROVISIONAL

function controlarBotonesExcluir2FA() {

  btnExcluir.classList.add('disabled');
  btnExcluir.disabled = true;
  btnReincluir.classList.add('disabled');
  btnReincluir.disabled = true;

  btnReiniciar2FA.classList.add('disabled');
  btnReiniciar2FA.disabled = true;

  usuarioSelect.addEventListener('change', function () {
    const selectedOption = usuarioSelect.value;
    const estado2FAItem = window.BDEstado2FA.find(item => item.Title === selectedOption);

    if (estado2FAItem) {
      if (estado2FAItem.ExcluirCuenta === null) {
        // Habilitar el botón "Excluir del MFA"
        btnExcluir.classList.remove('disabled');
        btnExcluir.disabled = false;

        // Deshabilitar el botón "Reincluir en MFA"
        btnReincluir.classList.add('disabled');
        btnReincluir.disabled = true;
      } else {
        // Habilitar el botón "Reincluir en MFA"
        btnReincluir.classList.remove('disabled');
        btnReincluir.disabled = false;

        // Deshabilitar el botón "Excluir del MFA"
        btnExcluir.classList.add('disabled');
        btnExcluir.disabled = true;
      }
    } else {
      // Si no se encuentra en BDEstado2FA, deshabilitar ambos botones
      btnExcluir.classList.add('disabled');
      btnExcluir.disabled = true;
      btnReincluir.classList.add('disabled');
      btnReincluir.disabled = true;
    }

    btnReiniciar2FA.classList.remove('disabled');
    btnReiniciar2FA.disabled = false;

  });
}

function mostrarUsuariosUltimos30Dias() {
  const tabla = document.querySelector("#tablaMovimientosLicencia tbody");
  const usuarios = window.usuariosEntraID || [];
  const licenciasPrincipales = Object.values(window.BDIdentificadorLicencias)
    .filter(licencia => licencia.LicenciaPrincipal === 1)
    .map(licencia => licencia.NombreLicencia);

  const fechaActual = new Date();
  const fechaLimite = new Date(fechaActual);
  fechaLimite.setDate(fechaActual.getDate() - 45);

  tabla.innerHTML = "";

  const usuariosFiltrados = usuarios.filter(usuario => {
    const fechaCreacion = usuario.createdDateTime;

    if (!fechaCreacion || fechaCreacion === "-") {
      return false;
    }

    const fechaUsuario = new Date(fechaCreacion.split(",")[0].trim().split("/").reverse().join("-"));
    return fechaUsuario >= fechaLimite;
  });

  usuariosFiltrados.sort((a, b) => {
    const fechaA = new Date(a.createdDateTime.split(",")[0].trim().split("/").reverse().join("-"));
    const fechaB = new Date(b.createdDateTime.split(",")[0].trim().split("/").reverse().join("-"));
    return fechaB - fechaA;
  });

  usuariosFiltrados.forEach(usuario => {
    const fila = document.createElement("tr");

    // Licencias formateadas
    const licencias = (usuario.assignedLicenses || "Sin Licencias").split(', ');
    const licenciasFormateadas = licencias.map(licencia => {
      if (licenciasPrincipales.includes(licencia)) {
        return `<strong>${licencia}</strong>`;
      }
      return licencia;
    }).join(', ');

    // Estilo para la fila según si tiene licencias o no
    if (usuario.assignedLicenses && usuario.assignedLicenses !== "Sin Licencia") {
      fila.style.backgroundColor = "#daf8d7"; // Verde
    } else {
      fila.style.backgroundColor = "#f8d7da"; // Rojo
    }

    // Crear fila
    fila.innerHTML = `
      <td>${usuario.createdDateTime || "Sin Información"}</td>
      <td>${usuario.displayName || "Sin Nombre"}</td>
      <td>${usuario.mail || "Sin Correo"}</td>
      <td>${usuario.jobTitle || "Sin Cargo"}</td>
      <td>${licenciasFormateadas}</td>
      <td>${usuario.officeLocation || "Sin Ubicación"}</td>
    `;

    tabla.appendChild(fila);
  });
}




async function manejarClic(event) {
  event.preventDefault(); // Previene la acción por defecto del botón

  const usuarioSelect = document.getElementById('usuarioSelect');
  const selectedOption = usuarioSelect.value;
  const estado2FAItem = window.usuariosEntraID.find(item => item.id === selectedOption);

  try {
    const result = await Swal.fire({
      title: 'Confirmación',
      html: `¿Estás seguro de reiniciar la configuración 2FA del usuario <strong>${estado2FAItem.mail}</strong>?`,
      icon: 'warning',
      showCancelButton: true,
      confirmButtonText: 'Sí, reiniciar',
      cancelButtonText: 'Cancelar'
    });


    if (result.isConfirmed) {

      const token = await obtenerToken();
      const removeResult = await removeMfaConfiguration(token, estado2FAItem.mail);

      if (removeResult.success) {
        mostrarAlerta('Éxito', removeResult.message, 'success');
      } else {
        mostrarAlerta('Error', removeResult.message, 'error');
      }
    }

  } catch (error) {
    mostrarAlerta('Error', 'Se produjo un error durante el proceso.', 'error');
  }
}

btnReiniciar2FA.addEventListener('click', manejarClic);


/* ---------------- Controlar los contenedores - Menú Principal -------------- */

document.addEventListener("DOMContentLoaded", function () {

    const menuLinks = document.querySelectorAll(".nav-menu a");

    menuLinks.forEach(link => {
        link.addEventListener("click", function (event) {
            event.preventDefault(); 

            const targetId = this.getAttribute("data-target");

            const allContents = document.querySelectorAll(".content");
            allContents.forEach(content => {
                content.classList.remove("active");
            });

            const targetContent = document.getElementById(targetId);
            if (targetContent) {
                targetContent.classList.add("active");
            }
        });
    });

    const firstContent = document.querySelector(".content");
    if (firstContent) {
        firstContent.classList.add("active");
    }
});




