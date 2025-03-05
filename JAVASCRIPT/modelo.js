/*--------------------------------------- Autenticarse en Microsoft Graph - Token de Acceso----------------------------------------------*/

window.usuariosEntraID = [];
window.estado2FA = [];

window.BDEstado2FA = [];
window.BDIdentificadorLicencias = []

let configuracionMsal = {};
let instanciaMsal;

/*---------------------------- Preparar sesión para consumir información de la API GRAPH----------------------------------------------------*/

async function buscarConfiguracionMsal() {
    const filter = "";
    const adicionales = '$select=ID,Title,authority,redirectUri';
    try {
        const datos = await consultarDatos(listaConfiguracionMsal, filter, adicionales);
        
        if (datos && datos.length > 0) {
            const configuracion = datos[0];

            configuracionMsal = {
                auth: {
                    clientId: configuracion.Title, 
                    authority: configuracion.authority,
                    redirectUri: configuracion.redirectUri
                }
            };
            return configuracionMsal; 
        } else {
            return null;
        }
    } catch (error) {
        mostrarAlerta('Error', 'Hubo un problema al iniciar sesión. Intenta nuevamente más tarde.', 'error');
        throw error;
    }
}

async function inicializarMsal() {
    try {
        await buscarConfiguracionMsal(); 
        instanciaMsal = new msal.PublicClientApplication(configuracionMsal); 
    } catch (error) {
        mostrarAlerta('Error', 'Hubo un problema al iniciar sesión. Intenta nuevamente más tarde.', 'error');
        throw error;
    }
}

async function iniciarSesion() {

    const solicitud = {
        scopes: ["user.read", "User.Read.All"]
    };

    try {
        const respuestaInicioSesion = await instanciaMsal.loginPopup(solicitud);
        instanciaMsal.setActiveAccount(respuestaInicioSesion.account);
        return respuestaInicioSesion.account;
    } catch (errorPopup) {
        if (errorPopup.errorCode === "popup_window_error" || errorPopup.errorCode === "empty_window_error") {
            console.warn("Ocurrió un error con la ventana emergente. Intentando método de redirección.");
            try {
                await instanciaMsal.loginRedirect(solicitud);
            } catch (errorRedireccion) {
                console.error('Error durante el inicio de sesión con redirección:', errorRedireccion);
                //  mostrarAlerta('Error', 'Hubo un problema al iniciar sesión. Intenta nuevamente más tarde.', 'error');
            }
        } else {
            console.error('Error durante el inicio de sesión:', errorPopup);
            // mostrarAlerta('Error', 'Hubo un problema al iniciar sesión. Intenta nuevamente más tarde.', 'error');
        }
    }
}


async function obtenerToken() {

    const cuentaActiva = instanciaMsal.getActiveAccount();

    if (!cuentaActiva) {
        await iniciarSesion();
    }

    const solicitud = {
        scopes: ["https://graph.microsoft.com/.default"],
        account: instanciaMsal.getActiveAccount()
    };

    try {
        const respuesta = await instanciaMsal.acquireTokenSilent(solicitud);
        return respuesta.accessToken;
    } catch (error) {
        if (error instanceof msal.InteractionRequiredAuthError) {
            return instanciaMsal.acquireTokenPopup(solicitud).then(respuesta => respuesta.accessToken)
                .catch(() => instanciaMsal.acquireTokenRedirect(solicitud));
        } else {
            console.error('Error al obtener el token en modo silencioso:', error);
            //mostrarAlerta('Error', 'Hubo un problema al obtener el token. Recargar la Página.', 'error');
        }
    }
}


async function inicializarAutenticacion() {

    try {
        const respuesta = await instanciaMsal.handleRedirectPromise();
        if (respuesta) {
            instanciaMsal.setActiveAccount(respuesta.account);
        } else {
            const cuentaActiva = instanciaMsal.getActiveAccount();
            if (!cuentaActiva) {
                await iniciarSesion();
            }
        }

        const token = await obtenerToken();

    } catch (error) {
        console.error('Error durante la inicialización de autenticación:', error);
        // mostrarAlerta('Error', 'Hubo un problema durante la inicialización de la autenticación. Intenta nuevamente más tarde.', 'error');
    }
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


/*--------------------------------------- Consultas a Microsoft Graph ----------------------------------------------*/


const userId = 'abogoya@icontec.onmicrosoft.com'; // O el ID del usuario
getAssignedLicensesForUser(userId);

async function getAssignedLicensesForUser(userId) {
    const token = await obtenerToken();
    let url = `https://graph.microsoft.com/v1.0/users/${userId}?$select=id,displayName,userPrincipalName,assignedLicenses`;

    const headers = {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
    };

    console.time('Tiempo de consulta de licencias del usuario');

    try {
        const response = await fetch(url, { headers });
        if (!response.ok) {
            throw new Error(`Error al obtener las licencias del usuario: ${response.statusText}`);
        }

        const userData = await response.json();

        console.log(`Licencias asignadas al usuario ${userData.displayName}:`, userData.assignedLicenses);

        console.timeEnd('Tiempo de consulta de licencias del usuario');

        return userData.assignedLicenses;

    } catch (error) {
        console.error(`Error al consultar las licencias del usuario: ${error.message}`);
        console.timeEnd('Tiempo de consulta de licencias del usuario');
    }
}



//const groupId = '24c16128-9549-47ae-8298-2906576e0b3a'; 
//getAllGroupMembers(groupId);

async function getAllGroupMembers(groupId) {
    const token = await obtenerToken();
    let url = `https://graph.microsoft.com/v1.0/groups/${groupId}/members?$top=999&$select=id,displayName,userPrincipalName`;

    const headers = {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
    };

    let members = [];
    let nextLink = null;

    // Iniciar el cronómetro para medir el tiempo
    console.time('Tiempo de consulta de miembros del grupo');

    try {
        do {
            const response = await fetch(url, { headers });
            if (!response.ok) {
                throw new Error(`Error al obtener los miembros del grupo: ${response.statusText}`);
            }

            const data = await response.json();
            members = members.concat(data.value);

            // Verificar si hay más páginas
            nextLink = data['@odata.nextLink'];
            url = nextLink;

        } while (nextLink);

        console.log("Total de miembros del grupo:", members.length);
        console.log("Datos de los miembros:", members);

        // Finalizar el cronómetro y mostrar el tiempo transcurrido
        console.timeEnd('Tiempo de consulta de miembros del grupo');

        return members;

    } catch (error) {
        console.error(`Error al consultar los miembros del grupo: ${error.message}`);
        console.timeEnd('Tiempo de consulta de miembros del grupo');
    }
}




/* Esta funcion se encarga de consultar todos los usuarios miembros de Entra ID, obtenie datos de usaurio y el Id de las licencias asignadas. */

//const userId = 'f5432724-812c-4472-bc94-e70a8c9255ce'; 
//getMfaStatus(userId);


//const url = `https://graph.microsoft.com/v1.0/users/${userId}/authentication/methods`;

async function getMfaStatus(userId) {
    const token = await obtenerToken();
    const url = `https://graph.microsoft.com/v1.0/users/${userId}/authentication/methods`;

    const headers = {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
    };

    try {
        const response = await fetch(url, { headers });
        if (!response.ok) {
            throw new Error(`Error al obtener los métodos de autenticación: ${response.statusText}`);
        }

        const data = await response.json();
        const methods = data.value;

        console.log(data);
        console.log("_______________________________________________________________");
        console.log(methods);

        const mfaStatus = {};

        methods.forEach(method => {

            switch (method.methodType) {
                case 'phone':
                    mfaStatus.phone = method.phoneNumber ? true : false;
                    break;
                case 'email':
                    mfaStatus.email = method.emailAddress ? true : false;
                    break;
                case 'app':
                    mfaStatus.app = method.appName ? true : false;
                    break;
                default:
                    mfaStatus.unknown = true;
                    break;
            }
        });

        console.log('Estado MFA:', mfaStatus);
        return mfaStatus;

    } catch (error) {
        console.error(`Error al consultar el estado de MFA: ${error.message}`);
    }
}


//---------------------------------------------------- Consultas a Graph ------------------------------------------------------------//


async function obtenerUsuarios(token, consultaConSignIn = false) {
    try {
        let datosUsuarios = [];

        const obtenerUsuariosPagina = async (url) => {
            const respuestaUsuarios = await fetch(url, {
                headers: {
                    'Authorization': `Bearer ${token}`
                }
            });

            if (!respuestaUsuarios.ok) {
                throw new Error(`Error al obtener usuarios: ${respuestaUsuarios.status} ${respuestaUsuarios.statusText}`);
            }

            const usuariosPagina = await respuestaUsuarios.json();
            datosUsuarios = datosUsuarios.concat(usuariosPagina.value);

            if (usuariosPagina['@odata.nextLink']) {
                return obtenerUsuariosPagina(usuariosPagina['@odata.nextLink']);
            }
        };

        const url = consultaConSignIn
            ? `https://graph.microsoft.com/v1.0/users?$filter=userType eq 'Member'&$select=id,displayName,mail,userPrincipalName,assignedLicenses,jobTitle,officeLocation,accountEnabled,createdDateTime,signInActivity`
            : `https://graph.microsoft.com/v1.0/users?$filter=userType eq 'Member'&$select=id,displayName,mail,userPrincipalName,assignedLicenses,jobTitle,officeLocation,accountEnabled,createdDateTime`;

        await obtenerUsuariosPagina(url);

        const usuariosConLicencias = datosUsuarios.map(usuario => {
            let licenciasPrincipales = [];
            let otrasLicencias = [];
            let skuIds = []; // Nuevo arreglo para almacenar los skuIds

            if (usuario.assignedLicenses && usuario.assignedLicenses.length > 0) {
                usuario.assignedLicenses.forEach(license => {

                    const skuId = license.skuId;

                    const licenciaNombre = window.BDIdentificadorLicencias[skuId]?.NombreLicencia || 'Licencia No Identificada';

                    skuIds.push(skuId);

                    if (window.BDIdentificadorLicencias[skuId]?.LicenciaPrincipal === 1) {
                        licenciasPrincipales.push(licenciaNombre);
                    } else {
                        otrasLicencias.push(licenciaNombre);
                    }
                });
            } else {
                otrasLicencias.push('Sin Licencia');
            }

            const assignedLicenses = [...licenciasPrincipales, ...otrasLicencias].join(', ');

            const lastSignInDateTime = consultaConSignIn
                ? (usuario.signInActivity && usuario.signInActivity.lastSignInDateTime
                    ? formatearFechaColombia(usuario.signInActivity.lastSignInDateTime)
                    : '-')
                : 'Sin Descargar';

            return {
                id: usuario.id || '-',
                displayName: usuario.displayName || '-',
                mail: usuario.userPrincipalName || usuario.mail || '-',
                jobTitle: usuario.jobTitle || '-',
                assignedLicenses: assignedLicenses,
                skuIds: skuIds.join(', '),
                officeLocation: usuario.officeLocation || '-',
                accountEnabled: usuario.accountEnabled ? 'Habilitado' : 'Bloqueado',
                createdDateTime: usuario.createdDateTime ? formatearFechaColombia(usuario.createdDateTime) : '-',
                lastSignInDateTime: lastSignInDateTime
            };
        });

        window.usuariosEntraID = usuariosConLicencias;
        generarLogin();
        return usuariosConLicencias;

    } catch (error) {
        console.error('Error al obtener usuarios y sus licencias:', error);
        mostrarAlerta('Error de Conexión', 'No se pudo establecer comunicación con el servidor. Por favor, recargue la página e intente nuevamente.', 'error');
    } finally {
        ocultarCargando();
    }
}



/* Esta función se encarga de extraer la cantidad de licencias adquiridas por cada licencia que tenemos. */

async function obtenerLicenciasDisponibles(token) {
    try {
        const respuestaLicencias = await fetch('https://graph.microsoft.com/v1.0/subscribedSkus', {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });

        if (!respuestaLicencias.ok) {
            throw new Error(`Error al obtener licencias: ${respuestaLicencias.status} ${respuestaLicencias.statusText}`);
        }

        const datosLicencias = await respuestaLicencias.json();
        const licenciasDisponibles = datosLicencias.value.map(licencia => ({
            nombre: window.BDIdentificadorLicencias[licencia.skuId]?.NombreLicencia || 'Licencia sin Identificar',
            total: licencia.prepaidUnits.enabled,
            licenciasExpiradas: licencia.prepaidUnits.warning,
            skuId: licencia.skuId
        }));

        return licenciasDisponibles;

    } catch (error) {
        console.error('Error al obtener licencias disponibles:', error);
    }
}



/* Esta función se encarga de extraer los registros de uditorias para luego extraer solo los movimientos de licencias.
Actualmente no se esta utilizando */

const consultarAuditoriasLogs = async (accessToken) => {
    try {
        const response = await fetch('https://graph.microsoft.com/v1.0/auditLogs/directoryAudits', {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });

        const data = await response.json();

        if (data.value && data.value.length > 0) {
            const licenseChangeEvents = data.value.filter(log => log.activityDisplayName === "Change user license");

            return licenseChangeEvents;

        } else {
            console.log("No se encontraron logs de auditoría.");
        }
    } catch (error) {
        console.error('Error al obtener los logs de auditoría:', error);
    }
};


function extraerInfoMovimientosLicencias(actividadesDeLicencia) {

    const informacionLicencias = [];

    if (actividadesDeLicencia && Array.isArray(actividadesDeLicencia)) {
        actividadesDeLicencia.forEach(actividad => {
            const informacion = {};

            const fechaHora = new Date(actividad.activityDateTime);
            const fechaHoraColombia = new Intl.DateTimeFormat('es-CO', {
                timeZone: 'America/Bogota',
                day: '2-digit',
                month: '2-digit',
                year: 'numeric',
                hour: 'numeric',
                minute: 'numeric',
                second: 'numeric'
            }).format(fechaHora);

            informacion.fechaHora = fechaHoraColombia;

            if (actividad.targetResources && actividad.targetResources.length > 0) {
                informacion.usuarioAfectado = actividad.targetResources[0].userPrincipalName;
            } else {
                informacion.usuarioAfectado = "System";
            }

            if (actividad.initiatedBy && actividad.initiatedBy.user && actividad.initiatedBy.user.userPrincipalName) {
                informacion.actor = actividad.initiatedBy.user.userPrincipalName;
            } else {
                informacion.actor = "System";
            }

            if (actividad.additionalDetails && actividad.additionalDetails[0] && actividad.additionalDetails[0].value) {
                informacion.id = actividad.additionalDetails[0].value;
            } else {
                informacion.id = "Sin información";
            }

            informacionLicencias.push(informacion);
        });
    } else {
        console.error('Error:', actividadesDeLicencia);
    }

    return informacionLicencias;
}



/* ----------------------------------------------------- FUNCIONES (CRUD), SHAREPOINT ------------------------------------------------ */

async function consultarDatos(lista, filter, adicionales) {
    const endpoint = `${urlSIME}/_api/web/lists/getbytitle('${lista}')/items?${adicionales}&${filter}`;
    let allResults = [];
    let nextEndpoint = endpoint;

    try {
        while (nextEndpoint) {
            const response = await fetch(nextEndpoint, {
                method: 'GET',
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'X-RequestDigest': document.getElementById("__REQUESTDIGEST").value,
                    'Content-Type': 'application/json;odata=verbose'
                }
            });

            if (!response.ok) {
                throw new Error(`Error al consultar la lista ${lista}: ${response.statusText}`);
            }

            const data = await response.json();
            allResults = allResults.concat(data.d.results);
            nextEndpoint = data.d.__next ? data.d.__next : null;
        }

        return allResults;

    } catch (error) {
        console.error('Error al consultar datos:', error);
        throw error;
    }
}



async function verificarExistencia(url) {
    try {
        const response = await fetch(url, {
            method: "GET",
            headers: {
                Accept: "application/json;odata=verbose"
            }
        });

        if (response.ok) {
            const data = await response.json();
            if (data.d.results.length > 0) {
                return data.d.results[0].ID;
            } else {
                return null;
            }
        } else {
            throw new Error(`Error HTTP: ${response.status}`);
        }
    } catch (error) {
        throw error;
    }
}


async function generarLogin() {
    try {
        const correoElectronico = await new Promise((resolve, reject) => {
            SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
                var clientContext = SP.ClientContext.get_current();
                var currentUser = clientContext.get_web().get_currentUser();
                clientContext.load(currentUser);
                clientContext.executeQueryAsync(
                    function () {
                        resolve(currentUser.get_email());
                    },
                    function (sender, args) {
                        reject('Error: ' + args.get_message());
                    }
                );
            });
        });

        const objetoDatos = {
            Title: "Inicio de sesion",
            Usuario: correoElectronico
        };

        const url = `${_spPageContextInfo.webAbsoluteUrl}/_api/web/lists/getbytitle('LogsAppSime')/items`;
        await insertarDatos(url, objetoDatos, 'LogsAppSime');

        const nombreLogin = window.usuariosEntraID.find(item => item.mail === correoElectronico)?.displayName;

        document.getElementById('loginNombre').textContent = nombreLogin;
        document.getElementById('loginCorreo').textContent = correoElectronico;

        const titulo = document.querySelector('.titulo h1');
        if (titulo) {
            titulo.style.marginLeft = '110px';
        }

    } catch (error) {
    }
}




async function insertarDatos(url, objetoDatos, nombreLista) {
    const spdata = `SP.Data.${nombreLista}ListItem`;
    objetoDatos.__metadata = { type: spdata };
    try {
        const response = await fetch(url, {
            method: "POST",
            headers: {
                Accept: "application/json;odata=verbose",
                "X-RequestDigest": document.getElementById("__REQUESTDIGEST").value,
                "Content-Type": "application/json;odata=verbose"
            },
            body: JSON.stringify(objetoDatos)
        });

        if (response.ok) {
            const data = await response.json();
            return data;
        } else {
            throw new Error(`Error HTTP: ${response.status}`);
        }
    } catch (error) {
        throw error;
    }
}


async function actualizarDatos(urlActualizar, objetoDatos, nombreLista) {
    const spdata = `SP.Data.${nombreLista}ListItem`;
    const dataActualizar = {
        ...objetoDatos,
        __metadata: { type: spdata }
    };

    try {
        const response = await fetch(urlActualizar, {
            method: "MERGE",
            headers: {
                Accept: "application/json;odata=verbose",
                "X-RequestDigest": document.getElementById("__REQUESTDIGEST").value,
                "Content-Type": "application/json;odata=verbose",
                "IF-MATCH": "*"
            },
            body: JSON.stringify(dataActualizar)
        });

        if (response.ok) {
            return true;
        } else {
            throw new Error(`Error HTTP: ${response.status}`);
        }
    } catch (error) {
        console.error("Error", error);
        throw error;
    }
}


async function obtenerUltimaModificacionLista(nombreLista) {
    const url = `${urlSIME}/_api/web/lists/getbytitle('${nombreLista}')/items?$orderby=Modified desc&$top=1&$select=Modified`;

    try {
        const response = await fetch(url, {
            method: 'GET',
            headers: {
                'Accept': 'application/json;odata=verbose'
            }
        });

        if (!response.ok) {
            throw new Error(`Error en la solicitud: ${response.statusText}`);
        }

        const data = await response.json();
        if (data.d.results.length > 0) {
            const ultimaModificacion = data.d.results[0].Modified;
            return ultimaModificacion;
        } else {
            return null;
        }
    } catch (error) {
        console.error('Error al obtener la última modificación:', error);
        throw error;
    }
}



async function removeMfaConfiguration(token, userEmail) {
    const urlGetMethods = `https://graph.microsoft.com/v1.0/users/${userEmail}/authentication/methods`;

    const options = {
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        }
    };

    try {

        const responseGet = await fetch(urlGetMethods, options);
        if (!responseGet.ok) {
            const errorData = await responseGet.json();
            return { success: false, message: `Failed to get authentication methods: ${errorData.error.message}` };
        }

        const methods = await responseGet.json();

        for (const method of methods.value) {
            const methodId = method.id;
            const urlDeleteMethod = `https://graph.microsoft.com/v1.0/users/${userEmail}/authentication/methods/${methodId}`;
            const responseDelete = await fetch(urlDeleteMethod, { ...options, method: 'DELETE' });

            if (!responseDelete.ok) {
                const errorData = await responseDelete.json();
                return { success: false, message: `Failed to remove MFA method ${methodId}: ${errorData.error.message}` };
            }
        }

        return { success: true, message: 'MFA configuration removed successfully.' };
    } catch (error) {
        console.error('Network error:', error);
        return { success: false, message: 'Network error occurred.' };
    }
}