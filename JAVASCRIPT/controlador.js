
const listaEstado2FA = "Estado2FA";
const listaIdentificadorLicencias ="IdentificadorLicencias";
const listaConfiguracionMsal = "ConfiguracionMsal";
let graficoIcontecOrg, graficoIcontecNet, graficoIcontecOtros, graficoIcontecTotal;


document.addEventListener("DOMContentLoaded", () => {
  
  const btnExcluirMFA = document.getElementById("btnExcluir");
  const btnIncluirMFA = document.getElementById("btnReincluir");
  const usuarioSelect = document.getElementById("usuarioSelect");

  btnExcluirMFA.addEventListener("click", (event) => {
    event.preventDefault();
    actualizarEstado2FA("1");
  });

  btnIncluirMFA.addEventListener("click", (event) => {
    event.preventDefault();
    actualizarEstado2FA(null);
  });


  async function actualizarEstado2FA(valorExcluirCuenta) {
    const usuarioID = usuarioSelect.value;
    const nombreLista = "Estado2FA";

    const urlVerificar = `${urlSIME}/_api/Web/Lists/GetByTitle('${nombreLista}')/items?$select=ID&$filter=Title eq '${usuarioID}'`;
    const urlBaseActualizar = `${urlSIME}/_api/Web/Lists/GetByTitle('${nombreLista}')/items`;

    const registroExiste = await verificarExistencia(urlVerificar);

    if (!registroExiste) {
        mostrarAlerta(
            "Error Inesperado",
            "Hubo un error inesperado. Por favor, contacte al administrador para obtener asistencia.",
            "error"
        );
        return; 
    }

    const urlActualizarConID = `${urlBaseActualizar}(${registroExiste})`;

    const objetoDatos = {
        ExcluirCuenta: valorExcluirCuenta
    };

    const usuarioSeleccionado = window.usuariosEntraID.find(usuario => usuario.id === usuarioID);
    const usuarioEmail = usuarioSeleccionado ? usuarioSeleccionado.mail : "desconocido";

    try {
        const resultado = await actualizarDatos(urlActualizarConID, objetoDatos, nombreLista);
        if (resultado) {
            Swal.fire({
                title: "Actualización Exitosa",
                text: `El usuario ${usuarioEmail} ha sido ${valorExcluirCuenta === "1" ? "excluido del" : "incluido en el"} MFA.`,
                icon: "success",
                confirmButtonColor: '#0459A1',
                confirmButtonText: 'Aceptar'
            }).then(() => {
                actualizarInformes2FA();
            });
        }
    } catch (error) {
        console.error("Error al actualizar el estado 2FA:", error);
        mostrarAlerta(
            "Error",
            "Hubo un problema al procesar la solicitud. Por favor, vuelva a intentarlo.",
            "error"
        );
    }
}

});


function configurarExclusionMFA() {
  const usuarioSelect = document.getElementById('usuarioSelect');
  const usuarioSearch = document.getElementById('usuarioSearch');

  function llenarDesplegableUsuarios() {
    usuarioSelect.innerHTML = '';

    const usuariosConLicencia = window.usuariosEntraID.filter(
      usuario => usuario.assignedLicenses && usuario.assignedLicenses !== "Sin Licencia"
    );

    usuariosConLicencia.forEach(usuario => {
      const option = document.createElement('option');
      option.value = usuario.id;
      option.textContent = `${usuario.displayName} (${usuario.mail})`;
      usuarioSelect.appendChild(option);
    });
  }

  usuarioSearch.addEventListener('input', function () {
    const filter = usuarioSearch.value.toLowerCase();
    Array.from(usuarioSelect.options).forEach(option => {
      const text = option.text.toLowerCase();
      option.style.display = text.includes(filter) ? '' : 'none';
    });
  });

  usuarioSearch.addEventListener('keydown', function (event) {
    if (event.key === 'Enter') {
      event.preventDefault(); 
      
    }
  });

  llenarDesplegableUsuarios();
}

document.getElementById('usuarioSearch').addEventListener('input', (event) => {
  controlarBotonesExcluir2FA();
  document.getElementById('usuarioSelect').selectedIndex = -1;
  
});

window.addEventListener('DOMContentLoaded', async function () {
  try {

    $('#tablaDatos').DataTable().destroy();

    mostrarCargando('Preparando datos, por favor espere...');

    await inicializarMsal();
    await inicializarAutenticacion();
    const token = await obtenerToken();

    await buscarDatosIdentificadorLicencias();

    const [datosUsuarios, licenciasTotales, datosEstado2FA, ultimaActualizacion2FA] = await Promise.all([
      obtenerUsuarios(token),
      obtenerLicenciasDisponibles(token),
      buscarDatosEstado2FA(),
      obtenerUltimaModificacionLista('Estado2FA')
    ]);

    const datosTabla = prepararDatosParaTabla(datosUsuarios, licenciasTotales);
    const usuariosConMFA = await generarObjetoUsuariosMFA(datosEstado2FA);

    showContainerInfo();
    mostrarLicenciasUsoTabla(datosTabla);
    generarOpcionesFiltroLicencias(datosUsuarios);
    mostrarUsuariosTabla(datosUsuarios);
    mostrarUsuarios2FATabla(usuariosConMFA);

    actualizarFechaUltimaModificacion(ultimaActualizacion2FA);

    const datosTotales = usuariosConMFA.icontecOrg
      .concat(usuariosConMFA.icontecNet)
      .concat(usuariosConMFA.otros)
      .concat(usuariosConMFA.excluidos);

    // Generar múltiples gráficos y almacenar sus instancias
    graficoIcontecOrg = generarGraficoTorta('grafico2FAIcontecOrg', usuariosConMFA.icontecOrg, graficoIcontecOrg);
    graficoIcontecNet = generarGraficoTorta('grafico2FAIcontecNet', usuariosConMFA.icontecNet, graficoIcontecNet);
    graficoIcontecOtros = generarGraficoTorta('grafico2FAIcontecOtros', usuariosConMFA.otros, graficoIcontecOtros);
    graficoIcontecTotal = generarGraficoTorta('grafico2FAIcontecTotal', datosTotales, graficoIcontecTotal);
    actualizarEstadisticas2FA(datosTotales, usuariosConMFA);

    configurarExclusionMFA();
    mostrarUsuariosUltimos30Dias();

    ocultarCargando();
    mostrarTarjetaExito("La descarga de datos ha finalizado con éxito.");

    // document.querySelectorAll('.containerInfo').forEach(function (div) {div.style.display = 'block';});
    // document.querySelector('.container2FA').style.display = 'none';

  } catch (error) {
    console.error('Error general:', error);
    ocultarCargando();
    }
});


async function actualizarFechaUltimaModificacion(ultimaModificacion) {
  try {
    if (ultimaModificacion) {
      const fechaFormateada = formatearFechaColombia(ultimaModificacion);
      document.getElementById('ultimaFechaActualizacion').textContent = fechaFormateada;
    }
  } catch (error) {
    console.error('Error:', error);
  }
}


function eventoMantenerLogin() {
  setInterval(() => {
    UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);
  }, 3000);
}


async function buscarDatosEstado2FA() {
  const filter = "";
  const adicionales = '$select=ID,Title,estado,ExcluirCuenta';

  try {

    const datos = await consultarDatos(listaEstado2FA, filter, adicionales);
    window.BDEstado2FA = datos;
    return datos;

  } catch (error) {
    console.error('Error al buscar datos:', error);
    throw error;
  }
}

async function buscarDatosIdentificadorLicencias() {
  const filter = "";
  const adicionales = '$select=ID,Title,LicenciaSkuId,LicenciaPrincipal,LicenciaDePago';

  try {
    const datos = await consultarDatos(listaIdentificadorLicencias, filter, adicionales);

    datos.forEach(licencia => {
      window.BDIdentificadorLicencias[licencia.LicenciaSkuId] = {
        ID: licencia.ID,
        NombreLicencia: licencia.Title,
        LicenciaPrincipal: licencia.LicenciaPrincipal,
        LicenciaDePago: licencia.LicenciaDePago
      };
    });

  } catch (error) {
    console.error('Error al buscar datos:', error);
    throw error;
  }
}


async function generarObjetoUsuariosMFA(datos2FA) {

  const usuariosExcluidos = [];
  const usuariosConMFA = [];

  window.usuariosEntraID.forEach(usuario => {

    const estado2FAItem = datos2FA.find(item => item.Title === usuario.id);
    const excluirCuenta = estado2FAItem ? estado2FAItem.ExcluirCuenta : null;

    if (excluirCuenta === "1") {

        // Crear objeto con los datos del usuario excluido y agregar a usuariosExcluidos
        usuariosExcluidos.push({
            displayName: usuario.displayName,
            mail: usuario.mail,
            jobTitle: usuario.jobTitle,
            accountEnabled: usuario.accountEnabled,
            assignedLicenses: usuario.assignedLicenses,
            estado2FA: estado2FAItem ? (estado2FAItem.estado === '1' ? 'Habilitado' : 'Deshabilitado') : 'Desconocido',
            ExcluirCuenta: excluirCuenta 
        });
    } else if (usuario.assignedLicenses && usuario.assignedLicenses !== "Sin Licencia") {
        // Si no está excluido y tiene licencias asignadas, agregar a usuariosConMFA
        usuariosConMFA.push({
            displayName: usuario.displayName,
            mail: usuario.mail,
            jobTitle: usuario.jobTitle,
            accountEnabled: usuario.accountEnabled,
            assignedLicenses: usuario.assignedLicenses,
            estado2FA: estado2FAItem ? (estado2FAItem.estado === '1' ? 'Habilitado' : 'Deshabilitado') : 'Desconocido',
            ExcluirCuenta: excluirCuenta 
        });
    }
});


  // Dividir usuarios con MFA en tres secciones según el dominio del correo electrónico
  const usuariosIcontecOrg = usuariosConMFA.filter(usuario => usuario.mail.endsWith("icontec.org"));
  const usuariosIcontecNet = usuariosConMFA.filter(usuario => usuario.mail.endsWith("icontec.net"));
  const usuariosOtros = usuariosConMFA.filter(usuario =>
    !usuario.mail.endsWith("icontec.org") && !usuario.mail.endsWith("icontec.net")
  );


  return {
    icontecOrg: usuariosIcontecOrg,
    icontecNet: usuariosIcontecNet,
    otros: usuariosOtros,
    excluidos: usuariosExcluidos
  };
}


function generarGraficoTorta(canvasId, datos, variableGrafico) {

  if (variableGrafico) {
    variableGrafico.destroy();
  }
  const habilitado = datos.filter(user => user.estado2FA === 'Habilitado').length;
  const deshabilitado = datos.filter(user => user.estado2FA === 'Deshabilitado').length;
  const ctx = document.getElementById(canvasId).getContext('2d');

  variableGrafico = new Chart(ctx, {
    type: 'pie',
    data: {
      labels: ['Habilitado', 'Deshabilitado'],
      datasets: [{
        data: [habilitado, deshabilitado],
        backgroundColor: [
          'rgba(102, 187, 106, 0.5)',
          'rgba(239, 83, 80, 0.5)'
        ],
        borderColor: [
          'rgba(56, 142, 60, 0.8)',
          'rgba(211, 47, 47, 0.8)'
        ],
        borderWidth: 2
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: {
          position: 'top',
          labels: {
            font: {
              size: 14,
              family: "'Helvetica Neue', 'Helvetica', 'Arial', sans-serif",
              weight: 'bold'
            },
            color: '#333'
          }
        },
        tooltip: {
          callbacks: {
            label: function (tooltipItem) {
              const value = tooltipItem.raw;
              const label = tooltipItem.label;
              return `${label}: ${value} usuarios`;
            }
          }
        },
        datalabels: {
          color: '#000',
          font: {
            size: 19,
            weight: 'bold'
          },
          formatter: (value, context) => {
            const total = context.chart._metasets[0].total;
            const percentage = (value / total * 100).toFixed(2);
            return `${percentage}%`;
          }
        }
      }
    },
    plugins: [ChartDataLabels]
  });

  return variableGrafico; // Retorna la instancia del gráfico para almacenarla
}


function generarGraficoLicencias(datos) {

  datos.sort((a, b) => b.disponibles - a.disponibles);

  const ctx = document.getElementById('graficoLicencias').getContext('2d');
  const labels = datos.map(item => item.licencia);

  const disponibles = datos.map(item => item.disponibles < 0 ? 0 : item.disponibles);

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

//Validar el archivo 2FA subido a la herramienta.
async function validarArchivo2FA() {

  ocultarResultado();

  const archivoSeleccionado = inputArchivo2FA.files.length > 0;
  const nombreArchivo = archivoSeleccionado ? inputArchivo2FA.files[0].name : null;

  if (archivoSeleccionado && nombreArchivo.endsWith(".csv")) {

    nombreArchivo2FA.textContent = nombreArchivo;
    nombreArchivo2FA.classList.remove("texto-normal");
    nombreArchivo2FA.classList.remove("texto-no-valido");
    nombreArchivo2FA.classList.add("texto-valido");
    botonCargarArchivo2FA.textContent = "Cambiar archivo";


    let datosArchivo2FA = await cargarArchivo2FA(inputArchivo2FA.files[0]);
    let resultadoVerificaciónDatos2FA = verificarIntegridadDatos(usuariosEntraID, datosArchivo2FA)

    if (resultadoVerificaciónDatos2FA) {

      botonIniciarActualizacion2FA.removeAttribute("disabled");
      botonIniciarActualizacion2FA.classList.remove("deshabilitado");
      botonIniciarActualizacion2FA.classList.add("habilitado");

      window.estado2FA = datosArchivo2FA;
      mostrarResultado(true);

    } else {
      botonIniciarActualizacion2FA.setAttribute("disabled", "disabled");
      botonIniciarActualizacion2FA.classList.add("deshabilitado");
      botonIniciarActualizacion2FA.classList.remove("habilitado");

      mostrarResultado(false);

    }

  } else if (!archivoSeleccionado) {
    nombreArchivo2FA.textContent = "No se seleccionó ningún archivo.";
    nombreArchivo2FA.classList.remove("texto-valido");
    nombreArchivo2FA.classList.add("texto-no-valido");
    botonCargarArchivo2FA.textContent = "Cargar archivo";

    botonIniciarActualizacion2FA.setAttribute("disabled", "disabled");
    botonIniciarActualizacion2FA.classList.add("deshabilitado");
    botonIniciarActualizacion2FA.classList.remove("habilitado");

  } else {
    nombreArchivo2FA.textContent = "El formato de archivo no es válido. Debe ser un archivo CSV.";
    nombreArchivo2FA.classList.remove("texto-valido");
    nombreArchivo2FA.classList.add("texto-no-valido");
    botonCargarArchivo2FA.textContent = "Cargar archivo";

    botonIniciarActualizacion2FA.setAttribute("disabled", "disabled");
    botonIniciarActualizacion2FA.classList.add("deshabilitado");
    botonIniciarActualizacion2FA.classList.remove("habilitado");
  }
}


async function cargarArchivo2FA(archivo2FA) {

  const reader = new FileReader();

  return new Promise((resolve, reject) => {
    reader.onload = (event) => {
      const text = event.target.result;
      const lines = text.split('\n');
      const result = [];

      lines.slice(1).forEach(line => {
        const [idUsuario, estado2FA] = line.split(',');
        if (idUsuario && estado2FA) {
          result.push({
            idUsuario: idUsuario.trim().replace(/"/g, ''),
            estado2FA: estado2FA.trim().replace(/"/g, '')
          });
        }
      });

      resolve(result);
    };

    reader.onerror = (error) => reject(error);
    reader.readAsText(archivo2FA);
  });
}


function verificarIntegridadDatos(usuariosEntraID, datosArchivo2FA) {

  const camposRequeridos = ['idUsuario', 'estado2FA'];
  const tieneCamposRequeridos = datosArchivo2FA.every(item =>
    camposRequeridos.every(campo => item.hasOwnProperty(campo))
  );

  if (!tieneCamposRequeridos) {
    return false;
  }

  const idsUsuariosEntraID = new Set(usuariosEntraID.map(usuario => usuario.id));
  const idsCoinciden = datosArchivo2FA.every(item => idsUsuariosEntraID.has(item.idUsuario));

  if (idsCoinciden) {
    return true;
  } else {
    return false;
  }
}

async function ejecutarActualizacion2FA() {
  const MAX_CONCURRENT_REQUESTS = 50; 
  let activeRequests = 0;

  botonIniciarActualizacion2FA.setAttribute("disabled", "disabled");
  botonIniciarActualizacion2FA.classList.add("deshabilitado");
  botonIniciarActualizacion2FA.classList.remove("habilitado");

  const totalRegistros = estado2FA.length;
  let registrosProcesados = 0;
  let registrosOK = 0;
  let registrosKO = 0;
  let registrosError429 = 0;

  const promesas = estado2FA.map((registro, index) => {
    return new Promise(async (resolve, reject) => {
      while (activeRequests >= MAX_CONCURRENT_REQUESTS) {
        await new Promise(r => setTimeout(r, 50)); // Espera a que haya espacio
      }
      activeRequests++;

      const estadoBinario = registro.estado2FA === "Habilitado" ? "1" : "0";
      const objetoDatos2FA = {
        Title: registro.idUsuario.toString(),
        estado: estadoBinario
      };

      const nombreListaEstado2FA = 'Estado2FA';
      const urlVerificar = `${urlSIME}/_api/Web/Lists/GetByTitle('${nombreListaEstado2FA}')/items?$select=ID&$filter=Title eq '${registro.idUsuario}'`;
      const urlBaseActualizar = `${urlSIME}/_api/Web/Lists/GetByTitle('${nombreListaEstado2FA}')/items`;
      const urlCrear = `${urlSIME}/_api/Web/Lists/GetByTitle('${nombreListaEstado2FA}')/items`;

      try {
        await procesarActualizacion(urlVerificar, urlBaseActualizar, urlCrear, objetoDatos2FA);
        registrosOK++;
      } catch (error) {
        if (error.message.includes("429")) {
          registrosError429++;
        } else {
          registrosKO++;
        }
      } finally {
        registrosProcesados++;
        activeRequests--;
        actualizarProgreso(registrosProcesados, totalRegistros);
        resolve();
      }
    });
  });

  try {
    await Promise.all(promesas);
  } catch (error) {
    console.error('Error durante la actualización:', error);
  }

  mostrarResultadoFinal(registrosOK, registrosKO, registrosError429, totalRegistros);
  actualizarInformes2FA();

}

async function actualizarInformes2FA() {

  mostrarCargando('Actualizando datos, por favor espere...');

  const [datosEstado2FA, ultimaActualizacion2FA] = await Promise.all([
    buscarDatosEstado2FA(),
    obtenerUltimaModificacionLista('Estado2FA')
  ]);


  const usuariosConMFA = await generarObjetoUsuariosMFA(datosEstado2FA);
  mostrarUsuarios2FATabla(usuariosConMFA);
  actualizarFechaUltimaModificacion(ultimaActualizacion2FA);

  const datosTotales = usuariosConMFA.icontecOrg
      .concat(usuariosConMFA.icontecNet)
      .concat(usuariosConMFA.otros)
      .concat(usuariosConMFA.excluidos);

  graficoIcontecOrg = generarGraficoTorta('grafico2FAIcontecOrg', usuariosConMFA.icontecOrg, graficoIcontecOrg);
  graficoIcontecNet = generarGraficoTorta('grafico2FAIcontecNet', usuariosConMFA.icontecNet, graficoIcontecNet);
  graficoIcontecOtros = generarGraficoTorta('grafico2FAIcontecOtros', usuariosConMFA.otros, graficoIcontecOtros);
  graficoIcontecTotal = generarGraficoTorta('grafico2FAIcontecTotal', datosTotales, graficoIcontecTotal);
  actualizarEstadisticas2FA(datosTotales, usuariosConMFA);

  configurarExclusionMFA();
  controlarBotonesExcluir2FA();
  document.getElementById('usuarioSelect').selectedIndex = -1;

  ocultarCargando();
  mostrarTarjetaExito("La actualización de datos ha finalizado.");


}

async function procesarActualizacion(urlVerificar, urlBaseActualizar, urlCrear, objetoDatos) {
  try {
    const registroExiste = await verificarExistencia(urlVerificar);

    if (registroExiste) {
      const urlActualizarConID = `${urlBaseActualizar}(${registroExiste})`;
      const dataActualizar = { estado: objetoDatos.estado };
      return await actualizarDatos(urlActualizarConID, dataActualizar, 'Estado2FA');
    } else {
      return await insertarDatos(urlCrear, objetoDatos, 'Estado2FA');
    }
  } catch (error) {
    console.error("Error al procesar la actualización:", error, "Datos del objeto:", objetoDatos);
    throw error;
  }
}


function actualizarProgreso(registrosProcesados, totalRegistros) {

  const porcentaje = Math.floor((registrosProcesados / totalRegistros) * 100);
  document.getElementById('mensajeResultado').innerText = `Registros procesados: ${registrosProcesados} de ${totalRegistros}`;
  document.getElementById('animacionResultado').innerText = `${porcentaje}%`;

}

function mostrarResultadoFinal(registrosOK, registrosKO, registrosError429, totalRegistros) {

  const mensajeResultado = document.getElementById('mensajeResultado');
  const animacionResultado = document.getElementById('animacionResultado');

  if (registrosError429 > 0 || registrosKO > 0) {
    mensajeResultado.innerHTML = `La actualización ha finalizado con Errores.<br><strong>Registros Actualizados: <span style="color: green;"><strong>${registrosOK}</strong></span><br>Registros, error al actualizar: <span style="color: red;"><strong>${registrosKO}</strong></span></strong>`;
    if (registrosError429 > 0) {
      mensajeResultado.innerHTML += `<br><span style="color: red;"><strong>${registrosError429}</strong></span> solicitudes fallaron debido al error 429 (Demasiadas solicitudes). Intente nuevamente más tarde.`;
    }
    animacionResultado.innerText = '✗';
    animacion.className = 'animacion error visible';
  } else {
    mensajeResultado.innerHTML = `
    La actualización ha finalizado con éxito.<br><strong>Registros actualizados: <span style="color: green;">${registrosOK}</span><br>Registros, error al actualizar: <span style="color: red;">${registrosKO}</span></strong>`;
    animacionResultado.innerText = '✓';
  }

  nombreArchivo2FA.textContent = "Ningún archivo seleccionado";
  nombreArchivo2FA.classList.remove("texto-valido", "texto-no-valido");
  nombreArchivo2FA.classList.add("texto-normal");
  botonCargarArchivo2FA.textContent = "Cargar archivo";
  inputArchivo2FA.value = '';

}


function contarLicencias(datosUsuarios) {

  const resultado = {};

  datosUsuarios.forEach(usuario => {
    
    const skuIds = usuario.skuIds.split(', ');

    skuIds.forEach(skuId => {
      const licenciaInfo = window.BDIdentificadorLicencias[skuId];
      if (licenciaInfo && licenciaInfo.LicenciaDePago === 1) {
        const nombreLicencia = licenciaInfo.NombreLicencia;

        if (resultado[nombreLicencia]) {
          resultado[nombreLicencia]++;
        } else {
          resultado[nombreLicencia] = 1;
        }
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
   
      const disponibles = licenciaTotal.total + (licenciaTotal.licenciasExpiradas || 0) - licenciaEnUso;

      return {
        licencia,
        disponibles,
        uso: licenciaEnUso,
        total: licenciaTotal.total,
        licenciasExpiradas: licenciaTotal.licenciasExpiradas || 0
      };
    } else {
      return {
        licencia,
        disponibles: 0,
        uso: licenciaEnUso,
        total: 0,
        licenciasExpiradas: 0
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

    const licenciaInfo = Object.values(window.BDIdentificadorLicencias).find(
      item => item.NombreLicencia === licencia
    );

    if (licenciaInfo && licenciaInfo.LicenciaDePago === 1) {
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


// SINCRONIZAR LA ÚLTIMA FECHA DE INICIO DE SESIÓN
document.getElementById('sincronizarFechaSesionBtn').addEventListener('click', async function (event) {

  deshabilitarBotones();
  event.preventDefault();

  const icono = document.getElementById('iconoSincronizarFechaSesion');
  const imagenOriginal = "../../JairoSevilla/SiteAssets/InformesEntraID/IMAGENES/descargar_info.png";
  const imagenRotar = "../../JairoSevilla/SiteAssets/InformesEntraID/IMAGENES/cargando.png";

  icono.src = imagenRotar;
  icono.classList.add('rotar');

  try {

    const token = await obtenerToken();
    const datosUsuarios = await obtenerUsuarios(token, true);

    mostrarUsuariosTabla(datosUsuarios);
    generarOpcionesFiltroLicencias(datosUsuarios);
    reiniciarFiltrosUsuarios()

    mostrarTarjetaExito("La descarga de datos ha finalizado con éxito.");

  } catch (error) {

    console.error('Error en la descarga de datos:', error);
    mostrarAlerta('Error', 'Hubo un problema al descargar los datos. Intenta nuevamente más tarde.', 'error');

  } finally {

    icono.src = imagenOriginal;
    icono.classList.remove('rotar');
    habilitarBotones();
  }

});

// SINCRONIZAR INFORMACIÓN BÁSICA DE LA TABLA
document.getElementById('sincronizarBtn').addEventListener('click', async function (event) {

  deshabilitarBotones();
  event.preventDefault();

  const icono = document.getElementById('iconoSincronizar');
  const imagenOriginal = "../../JairoSevilla/SiteAssets/InformesEntraID/IMAGENES/descargarII_info.png"; // Asume una imagen inicial para sincronización
  const imagenRotar = "../../JairoSevilla/SiteAssets/InformesEntraID/IMAGENES/cargando.png"; // Imagen que gira


  icono.src = imagenRotar;
  icono.classList.add('rotar');


  try {
    const token = await obtenerToken();
    const datosUsuarios = await obtenerUsuarios(token);
 
    generarOpcionesFiltroLicencias(datosUsuarios);
    mostrarUsuariosTabla(datosUsuarios);
    reiniciarFiltrosUsuarios()

    mostrarTarjetaExito("La sincronización de datos ha finalizado con éxito.");

  } catch (error) {
    console.error('Error en la sincronización:', error);
    mostrarAlerta('Error', 'Hubo un problema al sincronizar los datos. Intenta nuevamente más tarde.', 'error');
  } finally {

    icono.src = imagenOriginal;
    icono.classList.remove('rotar');
    habilitarBotones();

  }
});

//EXPORTAR LOS DATOS A EXCEL
document.getElementById('exportBtn').addEventListener('click', function (event) {
  event.preventDefault();

  var table = $('#tablaUsuarios').DataTable();
  var data = table.rows().data().toArray();

  mostrarTarjetaExito("Los registros se han exportado correctamente.");

  setTimeout(() => {
    exportarExcelDeTabla(data, 'Usuarios.xlsx');
  }, 500);

});


///////////////////////////// EXPORTAR TABLA USAURIOS A EXCEL /////////////////////////////////////////////

function separarFechaHora(fechaHoraStr) {
  if (fechaHoraStr === "-") {
    return {
      fecha: '',
      hora: ''
    };
  }

  const [fecha, hora] = fechaHoraStr.split(', ');
  return {
    fecha: fecha || '',
    hora: hora || ''
  };
}


function procesarLicenciasPrincipales(datosUsuarios) {
  return datosUsuarios.map(usuario => {
    const licenciasPrincipales = [];
    const otrasLicencias = [];

    usuario.assignedLicenses.split(', ').forEach(licenciaNombre => {
      const licenciaInfo = Object.values(window.BDIdentificadorLicencias).find(
        item => item.NombreLicencia === licenciaNombre
      );

      if (licenciaInfo?.LicenciaPrincipal === 1) {
        licenciasPrincipales.push(licenciaNombre);
      } else {
        otrasLicencias.push(licenciaNombre);
      }
    });

    return {
      ...usuario,
      assignedLicenses: otrasLicencias.length > 0 ? otrasLicencias.join(', ') : 'Sin Licencia',
      licenciaPrincipal: licenciasPrincipales.length > 0 ? licenciasPrincipales.join(', ') : 'Sin Licencia'
    };
  });
}

// function exportarExcelDeTabla(datosTabla) {
//   const datosProcesados = procesarLicenciasPrincipales(datosTabla);
//   var wb = XLSX.utils.book_new();

//   var datosFormateados = datosProcesados.map(function (row) {
//     const fechaHoraSeparadaCreacion = separarFechaHora(row.createdDateTime);
//     const fechaHoraSeparadaUltimoInicio = separarFechaHora(row.lastSignInDateTime);

//     return {
//       'NOMBRE COMPLETO': row.displayName === '-' ? '' : row.displayName,
//       'USUARIO': row.mail === '-' ? '' : row.mail,
//       'CARGO': row.jobTitle === '-' ? '' : row.jobTitle,
//       'OFICINA': row.officeLocation === '-' ? '' : row.officeLocation,
//       'LICENCIA PRINCIPAL': row.licenciaPrincipal === '-' ? '' : row.licenciaPrincipal,
//       'OTRAS LICENCIAS': row.assignedLicenses === '-' ? '' : row.assignedLicenses,
//       'INICIO DE SESIÓN': row.accountEnabled === '-' ? '' : row.accountEnabled,
//       'CREACIÓN (FECHA)': fechaHoraSeparadaCreacion.fecha,
//       'CREACIÓN (HORA)': fechaHoraSeparadaCreacion.hora,
//       'ÚLTIMO INICIO DE SESIÓN (FECHA)': fechaHoraSeparadaUltimoInicio.fecha,
//       'ÚLTIMO INICIO DE SESIÓN (HORA)': fechaHoraSeparadaUltimoInicio.hora
//     };
//   });

//   var ws = XLSX.utils.json_to_sheet(datosFormateados);

//   var colWidths = [
//     { wpx: 200 }, // Nombre Completo
//     { wpx: 180 }, // Usuario
//     { wpx: 150 }, // Cargo
//     { wpx: 150 }, // Ubicación
//     { wpx: 200 }, // Licencia Principal
//     { wpx: 200 }, // Otras Licencias
//     { wpx: 120 }, // Inicio de Sesión
//     { wpx: 100 }, // Creación (Fecha)
//     { wpx: 100 }, // Creación (Hora)
//     { wpx: 100 }, // Último Inicio de Sesión (Fecha)
//     { wpx: 100 }  // Último Inicio de Sesión (Hora)
//   ];
//   ws['!cols'] = colWidths;

//   XLSX.utils.book_append_sheet(wb, ws, 'Usuarios');

//   var now = new Date();
//   var year = now.getFullYear();
//   var month = (now.getMonth() + 1).toString().padStart(2, '0');
//   var day = now.getDate().toString().padStart(2, '0');
//   var hours = now.getHours().toString().padStart(2, '0');
//   var minutes = now.getMinutes().toString().padStart(2, '0');
//   var seconds = now.getSeconds().toString().padStart(2, '0');
//   var formattedDate = `${year}${month}${day}_${hours}${minutes}${seconds}`;
//   var nombreArchivo = 'Usuarios_SIME_' + formattedDate + '.xlsx';

//   XLSX.writeFile(wb, nombreArchivo);
// }


function exportarExcelDeTabla(datosTabla) {
  const datosProcesados = procesarLicenciasPrincipales(datosTabla);
  var wb = XLSX.utils.book_new();

  var datosFormateados = datosProcesados.map(function (row) {
    const fechaHoraSeparadaCreacion = separarFechaHora(row.createdDateTime);
    const fechaHoraSeparadaUltimoInicio = separarFechaHora(row.lastSignInDateTime);

    return {
      'NOMBRE COMPLETO': row.displayName === '-' ? '' : row.displayName,
      'USUARIO': row.mail === '-' ? '' : row.mail,
      'CARGO': row.jobTitle === '-' ? '' : row.jobTitle,
      'OFICINA': row.officeLocation === '-' ? '' : row.officeLocation,
      'LICENCIA PRINCIPAL': row.licenciaPrincipal === '-' ? '' : row.licenciaPrincipal,
      'OTRAS LICENCIAS': row.assignedLicenses === '-' ? '' : row.assignedLicenses,
      'INICIO DE SESIÓN': row.accountEnabled === '-' ? '' : row.accountEnabled,
      'CREACIÓN (FECHA)': fechaHoraSeparadaCreacion.fecha,
      'CREACIÓN (HORA)': fechaHoraSeparadaCreacion.hora,
      'ÚLTIMO INICIO DE SESIÓN (FECHA)': fechaHoraSeparadaUltimoInicio.fecha,
      'ÚLTIMO INICIO DE SESIÓN (HORA)': fechaHoraSeparadaUltimoInicio.hora
    };
  });

  // Convertir las fechas a formato de Excel
  var ws = XLSX.utils.json_to_sheet(datosFormateados);

  // Configuración de columnas para que las fechas sean detectadas como tales
  ws['!cols'] = [
    { wpx: 200 }, // Nombre Completo
    { wpx: 180 }, // Usuario
    { wpx: 150 }, // Cargo
    { wpx: 150 }, // Ubicación
    { wpx: 200 }, // Licencia Principal
    { wpx: 200 }, // Otras Licencias
    { wpx: 120 }, // Inicio de Sesión
    { wpx: 100 }, // Creación (Fecha)
    { wpx: 100 }, // Creación (Hora)
    { wpx: 100 }, // Último Inicio de Sesión (Fecha)
    { wpx: 100 }  // Último Inicio de Sesión (Hora)
  ];

  // Asegurarse de que las fechas estén en formato de Excel
  const dateColumns = ['CREACIÓN (FECHA)', 'ÚLTIMO INICIO DE SESIÓN (FECHA)'];

  dateColumns.forEach(function (col) {
    for (var rowIndex = 1; rowIndex <= datosFormateados.length; rowIndex++) {
      const cellRef = `${col}${rowIndex + 1}`; // Ajusta el índice para comenzar en la fila 2
      if (ws[cellRef]) {
        const cellValue = ws[cellRef].v;
        const excelDate = new Date(cellValue).getTime();
        if (!isNaN(excelDate)) {
          ws[cellRef].t = 'd'; // Formato de fecha
          ws[cellRef].v = excelDate / 1000 / 60 / 60 / 24 + 25569; // Convierte a formato de fecha de Excel
        }
      }
    }
  });

  // Agregar la hoja al libro
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
