/*--------------------------------------- Autenticarse en Microsoft Graph - Token de Acceso----------------------------------------------*/

const configuracionMsal = {
    auth: {
        clientId: "98888efe-4cb3-4bfc-a713-d6bfee0d592f",
        authority: "https://login.microsoftonline.com/ba76129b-6751-49e9-99a6-08b57c4e80fc",
        redirectUri: "https://icontec.sharepoint.com/sites/JairoSevilla/SitePages/InformesEntraID.aspx"
    }
};

const instanciaMsal = new msal.PublicClientApplication(configuracionMsal);

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
                mostrarAlerta('Error', 'Hubo un problema al iniciar sesión. Intenta nuevamente más tarde.', 'error');
            }
        } else {
            console.error('Error durante el inicio de sesión:', errorPopup);
            mostrarAlerta('Error', 'Hubo un problema al iniciar sesión. Intenta nuevamente más tarde.', 'error');
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
            mostrarAlerta('Error', 'Hubo un problema al obtener el token. Recargar la Página.', 'error');
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
        // console.log('Token de acceso:', token);

    } catch (error) {
        console.error('Error durante la inicialización de autenticación:', error);
        mostrarAlerta('Error', 'Hubo un problema durante la inicialización de la autenticación. Intenta nuevamente más tarde.', 'error');
    }
}

/*--------------------------------------- Consultas a Microsoft Graph ----------------------------------------------*/

const licenciasMap = {

    '05e9a617-0261-4cee-bb44-138d3ef5d965': 'Microsoft 365 E3', // SPE_E3
    '6fd2c87f-b296-42f0-b197-1e91e994b900': 'Office 365 E3', // ENTERPRISEPACK
    '4b585984-651b-448a-9e53-3b10f069cf7f': 'Office 365 F3', // DESKLESSPACK

    'bc946dac-7877-4271-b2f7-99d2db13cd2c': 'Prueba de Dynamics 365 Customer Voice', // FORMS_PRO
    'a403ebcc-fae0-4ca2-8c8c-7a907fd6c235': 'Microsoft Fabric (Gratis)', //POWER_BI_STANDARD
    '46102f44-d912-47e7-b0ca-1bd7b70ada3b' :'Project plan 3 (for Departments)', //PROJECT_PLAN3_DEPT
    'f30db892-07e9-47e9-837c-80727f46fd3d': 'Microsoft Power Automate (Gratis)', // FLOW_FREE
    'f8a1db68-be16-40ed-86d5-cb42ce701560': 'Power BI Pro', // POWER_BI_PRO
    '4a51bf65-409c-4a91-b845-1121b571cc9d': 'Power Automate per user plan', // FLOW_PER_USER
    '53818b1b-4a27-454b-8896-0dba576410e6': 'Project Plan 3', // PROJECTPROFESSIONAL


    '1f2f344a-700d-42c9-9427-5cea1d5d7ba6': 'Prueba de Microsoft Stream', // STREAM
    '99049c9c-6011-4908-bf17-15f496e6519d': 'OneDrive for Business (Plan 2)', // SHAREPOINTSTORAGE
    '606b54a9-78d8-4298-ad8b-df6ef4481c80': 'Power Virtual Agents Viral Trial', // CCIBOTS_PRIVPREV_VIRAL
    
    'dcb1a3ae-b33f-4487-846a-a640262fadf4': 'Microsoft Power Apps for Developer', // POWERAPPS_VIRAL
    '338148b6-1b11-4102-afb9-f92b6cdc0f8d': 'Dynamics 365 P1 Trial for Information Workers', // DYN365_ENTERPRISE_P1_IW
    'e0dfc8b9-9531-4ec8-94b4-9fec23b05fc8': 'Microsoft Teams Exploratory', // Microsoft_Teams_Exploratory_Dept
    'c1d032e0-5619-4761-9b5c-75b6831e1711': 'Power BI Premium por Usuario', // PBI_PREMIUM_PER_USER
    '3f9f06f5-3c31-472c-985f-62d9c10ec167': 'Power Pages vTrial for Makers', // Power_Pages_vTrial_for_Makers
    'b30411f5-fea1-4a59-9ad9-3db7c7ead579': 'Power Apps Premium', // POWERAPPS_PER_USER
   
    '52ea0e27-ae73-4983-a08f-13561ebdb823': 'Teams Premium (for Departments)', // Teams_Premium_(for_Departments)
    '4cde982a-ede4-4409-9ae6-b003453c8ea6': 'Salas de Microsoft Teams Pro', // Microsoft_Teams_Rooms_Pro
    '5b631642-bd26-49fe-bd20-1daaa972ef80': 'Microsoft Power Apps for Developer', // POWERAPPS_DEV
};


//////////////////////////////////////// VERIFICACIONES EXTRAS /////////////////////////////////

// alert("Ejecutar 2FA");
// verificarEstado2FA("jairo.sevilla@icontec.net", token);

// async function verificarEstado2FA(userId, token) {
//     const metodosAutenticacion = await consultarMetodosAutenticacion(userId, token);
//     const estado2FA = determinarEstado2FA(metodosAutenticacion);
//     console.log(`${userId}: ${estado2FA}`);
// }

// async function consultarMetodosAutenticacion(userId, token) {
//     const url = `https://graph.microsoft.com/v1.0/users/${userId}/authentication/methods`;

//     try {
//         const response = await fetch(url, {
//             method: 'GET',
//             headers: {
//                 'Authorization': `Bearer ${token}`,
//                 'Content-Type': 'application/json'
//             }
//         });

//         if (response.ok) {
//             const data = await response.json();
//             return data.value; 
//         } else {
//             console.error(`Error al consultar los métodos de autenticación para el usuario ${userId}:`, response.status, response.statusText);
//             return null; 
//         }
//     } catch (error) {
//         console.error(`Error al hacer la solicitud para el usuario ${userId}:`, error);
//         return null; 
//     }
// }

// function determinarEstado2FA(metodosAutenticacion) {
//     if (!metodosAutenticacion || metodosAutenticacion.length === 0) {
//         return "DESHABILITADO";
//     }
//     const tieneAuthenticator = metodosAutenticacion.some(metodo => 
//         metodo['@odata.type'] === '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod'
//     );
//     return tieneAuthenticator ? "HABILITADO" : "DESHABILITADO";
// }




// async function obtenerInfoUsuario(userName) {

//     const token = await obtenerToken();
//     // const url = `https://graph.microsoft.com/v1.0/users?$filter=mail eq '${userName}'&$select=displayName,mail,assignedLicenses`;
//     const url = `https://graph.microsoft.com/v1.0/users?$filter=mail eq '${userName}'`;

//     const respuestaUsuario = await fetch(url, {
//         headers: {
//             'Authorization': `Bearer ${token}`
//         }
//     });

//     if (!respuestaUsuario.ok) {
//         throw new Error(`Error al obtener usuario: ${respuestaUsuario.status} ${respuestaUsuario.statusText}`);
//     }

//     const datosUsuario = await respuestaUsuario.json();

//     if (datosUsuario.value.length === 0) {
//         console.log(`No se encontró el usuario con correo ${userName}`);
//         return;
//     }

//     const usuario = datosUsuario.value[0];
//     console.log('Información del usuario:', usuario);
// }


///////////////////////////////////////////////////////////////////////////////////////


async function obtenerUsuarios(token) {
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

        await obtenerUsuariosPagina(`https://graph.microsoft.com/v1.0/users?$filter=userType eq 'Member'&$select=displayName,mail,assignedLicenses,jobTitle,officeLocation,accountEnabled,createdDateTime`);

        const usuariosConLicencias = datosUsuarios.map(usuario => {
            let licenciasPrincipales = [];
            let otrasLicencias = [];

            if (usuario.assignedLicenses && usuario.assignedLicenses.length > 0) {
                usuario.assignedLicenses.forEach(license => {
                    const licenciaNombre = licenciasMap[license.skuId] || 'Licencia No Identificada';

                    if (['Microsoft 365 E3', 'Office 365 E3', 'Office 365 F3'].includes(licenciaNombre)) {
                        licenciasPrincipales.push(licenciaNombre);
                    } else {
                        otrasLicencias.push(licenciaNombre);
                    }
                });
            } else {
                otrasLicencias.push('Sin Licencia');
            }

            const assignedLicenses = [...licenciasPrincipales, ...otrasLicencias].join(', ');

            return {
                displayName: usuario.displayName || '-',
                mail: usuario.mail || '-',
                jobTitle: usuario.jobTitle || '-',
                assignedLicenses: assignedLicenses,
                officeLocation: usuario.officeLocation || '-',
                accountEnabled: usuario.accountEnabled ? 'Habilitado' : 'Bloqueado',
                createdDateTime: usuario.createdDateTime ? formatearFechaColombia(usuario.createdDateTime) : '-'
            };
        });

        return usuariosConLicencias;

    } catch (error) {
        console.error('Error al obtener usuarios y sus licencias:', error);
    } finally {
        ocultarCargando();
    }
}


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
            nombre: licenciasMap[licencia.skuId] || 'OTRA LICENCIA',
            total: licencia.prepaidUnits.enabled
        }));

        return licenciasDisponibles;

    } catch (error) {
        console.error('Error al obtener licencias disponibles:', error);
    }
}


const fetchAuditLogs = async (accessToken) => {
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


// function extraerInfoMovimientosLicencias(actividadesDeLicencia) {

//     console.log(actividadesDeLicencia);

//     const informacionLicencias = [];

//     if (actividadesDeLicencia && Array.isArray(actividadesDeLicencia)) {
//         actividadesDeLicencia.forEach(actividad => {
//             const informacion = {};

//             const fechaHora = new Date(actividad.activityDateTime);
//             const fechaHoraColombia = new Intl.DateTimeFormat('es-CO', {
//                 timeZone: 'America/Bogota',
//                 day: '2-digit',
//                 month: '2-digit',
//                 year: 'numeric',
//                 hour: 'numeric',
//                 minute: 'numeric',
//                 second: 'numeric'
//             }).format(fechaHora);

//             informacion.fechaHora = fechaHoraColombia;

//             if (actividad.targetResources && actividad.targetResources.length > 0) {
//                 informacion.usuarioAfectado = actividad.targetResources[0].userPrincipalName;
//             } else {
//                 informacion.usuarioAfectado = "System";
//             }

//             if (actividad.initiatedBy && actividad.initiatedBy.user && actividad.initiatedBy.user.userPrincipalName) {
//                 informacion.actor = actividad.initiatedBy.user.userPrincipalName;
//             } else {
//                 informacion.actor = "System";
//             }

//             informacion.id = Math.floor(Math.random() * 1000);

//             informacionLicencias.push(informacion);
//         });
//     } else {
//         console.error('Error:', actividadesDeLicencia);
//     }

//     console.log(informacionLicencias);
//     return informacionLicencias;
// }

function extraerInfoMovimientosLicencias(actividadesDeLicencia) {

    console.log(actividadesDeLicencia);

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

            if(actividad.additionalDetails && actividad.additionalDetails[0] && actividad.additionalDetails[0].value){
                informacion.id = actividad.additionalDetails[0].value;
            }else{
                informacion.id = "Sin información";
            }

            informacionLicencias.push(informacion);
        });
    } else {
        console.error('Error:', actividadesDeLicencia);
    }

    console.log(informacionLicencias);
    return informacionLicencias;
}



















