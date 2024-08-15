/*
<div class="containerFiltros">
                        <div class="card" style="width: 100%;">
                            <div class="card-body">
                                <h2 class="card-title">Filtro</h2>
                                <div class="form-group">
                                    <label for="opciones">Licencias:</label>
                                    <div id="grupo1">
                                        <label class="btn btn-secondary">
                                            <input type="radio" name="opcion"
                                                value="hoy"
                                                onchange="aplicarFiltro()">
                                            Office 1
                                        </label>
                                        <label class="btn btn-secondary">
                                            <input type="radio" name="opcion"
                                                value="estaSemana"
                                                onchange="aplicarFiltro()">
                                            Office 2
                                        </label>
                                        <label class="btn btn-secondary">
                                            <input type="radio" name="opcion"
                                                value="esteMes"
                                                onchange="aplicarFiltro()">
                                            Office 3
                                        </label>
                                        <label class="btn btn-secondary">
                                            <input type="radio" name="opcion"
                                                value="esteAnio"
                                                onchange="aplicarFiltro()">
                                            Office 4
                                        </label>
                                        <label class="btn btn-secondary">
                                            <input type="radio" name="opcion"
                                                value="ultimos30Dias"
                                                onchange="aplicarFiltro()">
                                            Office 5
                                        </label>
                                        <label class="btn btn-secondary">
                                            <input type="radio" name="opcion"
                                                value="ultimos90Dias"
                                                onchange="aplicarFiltro()">
                                            Office 6
                                        </label>
                                        <label class="btn btn-secondary">
                                            <input type="radio" name="opcion"
                                                value="ultimoAnio"
                                                onchange="aplicarFiltro()">
                                            Office 7
                                        </label>
                                    </div> <!-- Cierre de grupo1 -->
                                </div> <!-- Cierre de form-group -->
                            </div> <!-- Cierre de card-body -->
                        </div> <!-- Cierre de card -->
                    </div> <!-- Cierre de containerFiltros -->

/*


async function obtenerLicenciasDeUsuario(correoUsuario) {
  try {
    const token = await obtenerToken();
    
    const obtenerUsuario = async (correo) => {
      const respuestaUsuario = await fetch(`https://graph.microsoft.com/v1.0/users/${correo}?$select=id,displayName,mail,assignedLicenses`, {
        headers: {
          'Authorization': `Bearer ${token}`
        }
      });

      if (!respuestaUsuario.ok) {
        throw new Error(`Error al obtener usuario: ${respuestaUsuario.status} ${respuestaUsuario.statusText}`);
      }

      return await respuestaUsuario.json();
    };

    const usuario = await obtenerUsuario(correoUsuario);

    console.log(usuario);

    if (!usuario) {
      console.log(`No se encontró el usuario con correo ${correoUsuario}`);
      return;
    }

    let licencias = [];

    if (usuario.assignedLicenses && usuario.assignedLicenses.length > 0) {
      for (const license of usuario.assignedLicenses) {
        if (license.skuId === '05e9a617-0261-4cee-bb44-138d3ef5d965') {
          licencias.push('Microsoft 365 E3');
        } else if (license.skuId === '18181a46-0d4e-45cd-891e-60aabd171b4e') {
          licencias.push('Office 365 E3');
        } else if (license.skuId === '8c7d2df8-86f0-4902-b2ed-a0458298f3b3') {
          licencias.push('Office 365 F3');
        } else {
          licencias.push('Otra licencia');
        }
      }
    } else {
      licencias.push('Sin licencia');
    }

    console.log(`Licencias para ${usuario.displayName} (${usuario.mail}): ${licencias.join(', ')}`);
  } catch (error) {
    console.error('Error al obtener licencias del usuario:', error);
  }
}

// Llamar a la función con el correo electrónico del usuario específico
obtenerLicenciasDeUsuario('diego.mendez@icontec.net');




async function getTotalLicenses() {
  try {
    console.log('Iniciando obtención y cálculo de licencias...');
    
    const token = await obtenerToken();

    if (!token) {
      throw new Error('No se pudo obtener el token de acceso.');
    }

    const licenseCounts = {
      'Microsoft 365 E3': 0,
      'Office 365 E3': 0,
      'Office 365 F3': 0
    };

    // Función para obtener todos los usuarios miembros
    const obtenerUsuariosMiembros = async () => {
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
          console.log('Paginación');
          return obtenerUsuariosPagina(usuariosPagina['@odata.nextLink']);
        }
      };

      await obtenerUsuariosPagina('https://graph.microsoft.com/v1.0/users?$filter=userType eq \'Member\'');
      return datosUsuarios;
    };

    // Obtener todos los usuarios miembros
    const usersData = await obtenerUsuariosMiembros();
    console.log(usersData);

    // Obtener detalles de licencia para cada usuario
    for (const user of usersData) {
      console.log(user);
      const licenseResponse = await fetch(`https://graph.microsoft.com/v1.0/users/${user.id}/licenseDetails`, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${token}`
        }
      });

      if (!licenseResponse.ok) {
        throw new Error(`Failed to fetch license info for user ${user.mail}: ${licenseResponse.status} ${licenseResponse.statusText}`);
      }

      const licenseInfo = await licenseResponse.json();

      for (const license of licenseInfo.value) {
        if (license.skuPartNumber === 'SPE_E3') {
          licenseCounts['Microsoft 365 E3']++;
        } else if (license.skuPartNumber === 'ENTERPRISEPACK') {
          licenseCounts['Office 365 E3']++;
        } else if (license.skuPartNumber === 'DESKLESSPACK') {
          licenseCounts['Office 365 F3']++;
        }
      }
    }

    // Imprimir la información en la consola
    console.log('Licencias:');
    console.log(`Microsoft 365 E3: ${licenseCounts['Microsoft 365 E3']}`);
    console.log(`Office 365 E3: ${licenseCounts['Office 365 E3']}`);
    console.log(`Office 365 F3: ${licenseCounts['Office 365 F3']}`);

  } catch (error) {
    console.error('Error fetching total licenses:', error);
  }
}



async function getUsersInDomain() {
  try {
    const token = await obtenerToken();

    // Encuentra todos los usuarios que sean miembros (usuarios locales del directorio)
    const usersResponse = await fetch(`https://graph.microsoft.com/v1.0/users?$filter=userType eq 'Member'`, {
      headers: {
        'Authorization': `Bearer ${token}`
      }
    });

    if (!usersResponse.ok) {
      throw new Error(`Failed to fetch users: ${usersResponse.status} ${usersResponse.statusText}`);
    }

    const usersData = await usersResponse.json();
    console.log(usersData);
  } catch (error) {
    console.error('Error fetching users:', error);
  }
}



async function getUsersInDomain() {
    try {
      const token = await obtenerToken();
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
  
      await obtenerUsuariosPagina(`https://graph.microsoft.com/v1.0/users?$filter=userType eq 'Member'`);
      mostrarDatosEnTabla(datosUsuarios);
  
    } catch (error) {
      console.error('Error al obtener usuarios:', error);
    }
  }
  
  
  
  function mostrarDatosEnTabla(datos) {
    $('#tablaDatos').DataTable().destroy();
    $('#tablaDatos tbody').empty();
  
    const columnas = [
      { data: 'id', title: 'ID' },
      { data: 'displayName', title: 'Nombre Completo' },
      { data: 'mail', title: 'Correo Electrónico' },
      { data: 'jobTitle', title: 'Puesto de Trabajo' }
    ];
  
    $('#tablaDatos').DataTable({
      data: datos,
      columns: columnas,
      language: {
        url: '//cdn.datatables.net/plug-ins/1.10.21/i18n/Spanish.json'
      }
    });
  }
  
  window.addEventListener('DOMContentLoaded', function () {
    getUsersInDomain();
  });


  const msalConfig = {
    auth: {
        clientId: "98888efe-4cb3-4bfc-a713-d6bfee0d592f",
        authority: "https://login.microsoftonline.com/ba76129b-6751-49e9-99a6-08b57c4e80fc",
        redirectUri: "https://icontec.sharepoint.com/sites/JairoSevilla/SitePages/microsoftGraph.aspx"
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

async function signIn() {
    const request = {
        scopes: ["user.read", "User.Read.All"]
    };

    try {
        const loginResponse = await msalInstance.loginPopup(request);
        msalInstance.setActiveAccount(loginResponse.account);
        return loginResponse.account;
    } catch (popupError) {
        if (popupError.errorCode === "popup_window_error" || popupError.errorCode === "empty_window_error") {
            console.warn("Popup error occurred. Trying redirect method.");
            try {
                await msalInstance.loginRedirect(request);
            } catch (redirectError) {
                console.error('Error during login with redirect:', redirectError);
            }
        } else {
            console.error('Error during login:', popupError);
        }
    }
}

async function getToken() {
    const activeAccount = msalInstance.getActiveAccount();

    if (!activeAccount) {
        await signIn();
    }

    const request = {
        scopes: ["https://graph.microsoft.com/.default"],
        account: msalInstance.getActiveAccount()
    };

    try {
        const response = await msalInstance.acquireTokenSilent(request);
        return response.accessToken;
    } catch (error) {
        if (error instanceof msal.InteractionRequiredAuthError) {
            return msalInstance.acquireTokenPopup(request).then(response => response.accessToken)
                .catch(() => msalInstance.acquireTokenRedirect(request));
        } else {
            console.error('Error acquiring token silently:', error);
        }
    }
}

async function getUserInfo(username) {
    try {
        const token = await getToken();

        if (!token) {
            throw new Error('No se pudo obtener el token de acceso.');
        }

        // Obtener información del usuario
        const userInfoResponse = await fetch(`https://graph.microsoft.com/v1.0/users/${username}`, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });

        if (!userInfoResponse.ok) {
            throw new Error(`Failed to fetch user info: ${userInfoResponse.status} ${userInfoResponse.statusText}`);
        }

        const userInfo = await userInfoResponse.json();

        // Obtener información de la licencia del usuario
        const licenseResponse = await fetch(`https://graph.microsoft.com/v1.0/users/${username}/licenseDetails`, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });

        if (!licenseResponse.ok) {
            throw new Error(`Failed to fetch license info: ${licenseResponse.status} ${licenseResponse.statusText}`);
        }

        const licenseInfo = await licenseResponse.json();

        // Obtener información de métodos de autenticación (MFA)
        const mfaResponse = await fetch(`https://graph.microsoft.com/beta/reports/credentialUserRegistrationDetails?$filter=userPrincipalName eq '${username}'`, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });

        if (!mfaResponse.ok) {
            throw new Error(`Failed to fetch MFA info: ${mfaResponse.status} ${mfaResponse.statusText}`);
        }

        const mfaInfo = await mfaResponse.json();
        console.log(mfaInfo);

        const mfaEnabled = mfaInfo.value.some(method => method.isMfaRegistered);
        console.log(mfaEnabled);

        // Filtrar y encontrar la licencia principal
        let mainLicense = 'No se encontró una licencia principal';

        for (const license of licenseInfo.value) {
            if (license.skuPartNumber === 'SPE_E3') {
                mainLicense = 'Microsoft 365 E3';
                break;
            } else if (license.skuPartNumber === 'ENTERPRISEPACK') {
                mainLicense = 'Office 365 E3';
                break;
            } else if (license.skuPartNumber === 'DESKLESSPACK') {
                mainLicense = 'Office 365 F3';
                break;
            }
        }

        // Mostrar la información en el HTML
        const userInfoDiv = document.getElementById("user-info");
        userInfoDiv.innerHTML = '';

        const infoItems = [
            { label: 'Nombre', value: userInfo.givenName },
            { label: 'Apellido', value: userInfo.surname },
            { label: 'Usuario', value: userInfo.mail },
            { label: 'Cargo', value: userInfo.jobTitle },
            { label: 'Oficina', value: userInfo.officeLocation },
            { label: 'Licencia', value: mainLicense },
            { label: 'MFA Habilitado', value: mfaEnabled ? 'Habilitado' : 'Deshabilitado' }
        ];

        infoItems.forEach(item => {
            const p = document.createElement('p');
            p.innerHTML = `<strong>${item.label}:</strong> ${item.value}`;
            userInfoDiv.appendChild(p);
        });

    } catch (error) {
        console.error('Error fetching user info:', error);
        document.getElementById("user-info").innerHTML = `<p class="error">Error: ${error.message}</p>`;
    }
}

async function getTotalLicenses() {
    try {
        const token = await getToken();

        if (!token) {
            throw new Error('No se pudo obtener el token de acceso.');
        }

        const licenseCounts = {
            'Microsoft 365 E3': 0,
            'Office 365 E3': 0,
            'Office 365 F3': 0
        };

        // Obtener información de todos los usuarios
        const usersResponse = await fetch('https://graph.microsoft.com/v1.0/users?$select=id,mail', {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });

        if (!usersResponse.ok) {
            throw new Error(`Failed to fetch users info: ${usersResponse.status} ${usersResponse.statusText}`);
        }

        const usersData = await usersResponse.json();
        console.log("Data de Usuarios");
        console.log(usersData);

        for (const user of usersData.value) {
            const licenseResponse = await fetch(`https://graph.microsoft.com/v1.0/users/${user.id}/licenseDetails`, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${token}`
                }
            });

            if (!licenseResponse.ok) {
                throw new Error(`Failed to fetch license info for user ${user.mail}: ${licenseResponse.status} ${licenseResponse.statusText}`);
            }

            const licenseInfo = await licenseResponse.json();

            for (const license of licenseInfo.value) {
                if (license.skuPartNumber === 'SPE_E3') {
                    licenseCounts['Microsoft 365 E3']++;
                } else if (license.skuPartNumber === 'ENTERPRISEPACK') {
                    licenseCounts['Office 365 E3']++;
                } else if (license.skuPartNumber === 'DESKLESSPACK') {
                    licenseCounts['Office 365 F3']++;
                }
            }
        }

        // Mostrar la información en el HTML
        const licenseInfoDiv = document.getElementById("license-info");
        licenseInfoDiv.innerHTML = '';

        const licenseItems = [
            { label: 'Microsoft 365 E3', value: licenseCounts['Microsoft 365 E3'] },
            { label: 'Office 365 E3', value: licenseCounts['Office 365 E3'] },
            { label: 'Office 365 F3', value: licenseCounts['Office 365 F3'] }
        ];

        licenseItems.forEach(item => {
            const p = document.createElement('p');
            p.innerHTML = `<strong>${item.label}:</strong> ${item.value}`;
            licenseInfoDiv.appendChild(p);
        });

    } catch (error) {
        console.error('Error fetching total licenses:', error);
        document.getElementById("license-info").innerHTML = `<p class="error">Error: ${error.message}</p>`;
    }
}

document.getElementById("search-button").addEventListener("click", function(event) {
    event.preventDefault();
    const username = document.getElementById("username-input").value;
    if (username) {
        getUserInfo(username);
    } else {
        alert("Por favor ingrese un correo electrónico.");
    }
});

document.getElementById("license-button").addEventListener("click", function(event) {
    event.preventDefault();
    getTotalLicenses();
});

msalInstance.handleRedirectPromise().then((response) => {
    if (response) {
        msalInstance.setActiveAccount(response.account);
    }
}).catch((error) => {
    console.error('Error handling redirect:', error);
});
  
  */




/*

window.addEventListener('DOMContentLoaded', function () {
  obtenerToken().then(token => {
    //obtenerEstado2FA(token);
    consultarUsuario(token, 'jortiza@icontec.org');
    consultarUsuario2FA(token, 'jortiza@icontec.org');
  }).catch(error => {
    console.error('Error al obtener el token:', error);
  });
});





async function consultarUsuario(token, userPrincipalName) {
  const url = `https://graph.microsoft.com/v1.0/users/${userPrincipalName}`;

  try {
    const respuesta = await fetch(url, {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    });

    if (!respuesta.ok) {
      throw new Error(`Error al consultar usuario: ${respuesta.status} ${respuesta.statusText}`);
    }

    const datosUsuario = await respuesta.json();
    console.log(datosUsuario);

  } catch (error) {
    console.error(`Error al consultar usuario ${userPrincipalName}:`, error);
    throw error;
  }
}

//----------------------------------------------------------------------------------------------------------

async function obtenerEstado2FA(token) {
  mostrarCargando('Obteniendo datos de Microsoft Graph API');

  try {
    let datosUsuarios = [];
    const maxIntentos = 5;

    const obtenerPagina = async (url, intentos = 0) => {
      try {
        const respuesta = await fetch(url, {
          headers: {
            'Authorization': `Bearer ${token}`
          }
        });

        if (respuesta.status === 429) {
          const retryAfterHeader = respuesta.headers.get('Retry-After');
          let retryAfterSeconds = 5;

          if (retryAfterHeader) {
            retryAfterSeconds = parseInt(retryAfterHeader);
          }

          mostrarCargando(`Solicitud limitada (429). Reintentando después de ${retryAfterSeconds} segundos...`);
          await new Promise(resolve => setTimeout(resolve, retryAfterSeconds * 1000));

          if (intentos < maxIntentos) {
            return obtenerPagina(url, intentos + 1);
          } else {
            throw new Error(`Se superó el número máximo de intentos (${maxIntentos}) para obtener el estado de 2FA.`);
          }
        } else if (!respuesta.ok) {
          throw new Error(`Error al obtener estado de 2FA: ${respuesta.status} ${respuesta.statusText}`);
        }

        const datos = await respuesta.json();
        console.log(datos);

        for (let usuario of datos.value) {
          if (usuario.userPrincipalName && 
              (usuario.userPrincipalName.endsWith('@icontec.net') ||
               usuario.userPrincipalName.endsWith('@la.icontec.org') ||
               usuario.userPrincipalName.endsWith('@icontec.org'))) {
            
            let estado2FA = 'Deshabilitado';

            if (usuario.isEnabled && usuario.isMfaRegistered && usuario.isRegistered && usuario.isCapable) {
              estado2FA = 'Habilitado';
            }
            
            datosUsuarios.push({
              id: usuario.id,
              displayName: usuario.userDisplayName,
              mail: usuario.userPrincipalName,
              jobTitle: 'No Disponible', 
              estado2FA: estado2FA
            });
          }
        }

        if (datos['@odata.nextLink']) {
          await obtenerPagina(datos['@odata.nextLink']);
        }
      } catch (error) {
        console.error('Error en solicitud:', error);
        if (intentos < maxIntentos) {
          mostrarCargando(`Reintentando en ${(intentos + 1) * 5} segundos...`);
          await new Promise(resolve => setTimeout(resolve, (intentos + 1) * 5000)); // Espera incremental
          return obtenerPagina(url, intentos + 1);
        } else {
          throw error;
        }
      }
    };

    await obtenerPagina(`https://graph.microsoft.com/beta/reports/credentialUserRegistrationDetails`);
    
    console.log("Usuarios extraidos");
    console.log(datosUsuarios);
    mostrarDatosEnTabla(datosUsuarios);
    ocultarCargando();

  } catch (error) {
    console.error('Error al obtener usuarios:', error);
    ocultarCargando();
  }
}


function mostrarDatosEnTabla(datos) {

  const tabla = $('#tablaDatos').DataTable();
  tabla.clear().destroy();
  
  $('#tablaDatos tbody').empty();
  $('#tablaDatos').DataTable({
    data: datos,
    columns: [
      { data: 'id', title: 'ID' },
      { data: 'displayName', title: 'Nombre Completo' },
      { data: 'mail', title: 'Correo Electrónico' },
      { data: 'jobTitle', title: 'Puesto de Trabajo' },
      { data: 'estado2FA', title: 'Estado 2FA' }
    ],
    language: {
      url: '//cdn.datatables.net/plug-ins/1.10.21/i18n/Spanish.json'
    }
  });
  ocultarCargando();
}

/*
async function obtenerEstado2FA(token) {
  const userPrincipalName = 'jairo.sevilla@icontec.net';
  console.log(userPrincipalName)
  const url = `https://graph.microsoft.com/beta/reports/credentialUserRegistrationDetails`;
  let intentos = 0;
  const maxIntentos = 5;

  while (intentos < maxIntentos) {
    try {
      const respuesta = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`
        }
      });

      if (respuesta.status === 429) {
        const retryAfterHeader = respuesta.headers.get('Retry-After');
        let retryAfterSeconds = 10; // Valor por defecto si no se especifica Retry-After

        if (retryAfterHeader) {
          retryAfterSeconds = parseInt(retryAfterHeader);
        }

        console.warn(`Solicitud limitada (429). Reintentando después de ${retryAfterSeconds} segundos.`);
        await new Promise(resolve => setTimeout(resolve, retryAfterSeconds * 1000));
        intentos++;
      } else if (!respuesta.ok) {
        throw new Error(`Error al obtener estado de 2FA: ${respuesta.status} ${respuesta.statusText}`);
      } else {
        const datos = await respuesta.json();
        console.log(datos)
        const usuario = datos.value.find(u => u.userPrincipalName === userPrincipalName);

        if (!usuario) {
          return 'Deshabilitado';
        }

        const estado2FA = usuario.authMethods.includes('phone') ? 'Habilitado' : 'Forzado';
        console.log(estado2FA);
        return estado2FA;
      }
    } catch (error) {
      if (intentos >= maxIntentos - 1) {
        throw error;
      } else {
        console.error('Error en solicitud:', error);
        intentos++;
        await new Promise(resolve => setTimeout(resolve, (intentos + 1) * 1000));
      }
    }
  }

  throw new Error(`Se superó el número máximo de intentos (${maxIntentos}) para obtener el estado de 2FA.`);
}


/*

async function getUsersInDomain() {
  try {
    const token = await obtenerToken();
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
      for (let usuario of usuariosPagina.value) {
        // Verificar si usuario.mail existe y termina con el dominio deseado
        if (usuario.mail && (usuario.mail.endsWith('@icontec.org') || usuario.mail.endsWith('@la.icontec.org'))) {
          usuario.estado2FA = await obtenerEstado2FA(usuario.userPrincipalName, token);
          datosUsuarios.push(usuario);
        }
      }

      if (usuariosPagina['@odata.nextLink']) {
        return obtenerUsuariosPagina(usuariosPagina['@odata.nextLink']);
      }
    };

    await obtenerUsuariosPagina(`https://graph.microsoft.com/v1.0/users?$filter=userType eq 'Member'`);
    mostrarDatosEnTabla(datosUsuarios);

  } catch (error) {
    console.error('Error al obtener usuarios:', error);
  }
}



async function obtenerEstado2FA(userPrincipalName, token) {
  console.log(userPrincipalName)
  const url = `https://graph.microsoft.com/beta/reports/credentialUserRegistrationDetails`;
  let intentos = 0;
  const maxIntentos = 5;

  while (intentos < maxIntentos) {
    try {
      const respuesta = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`
        }
      });

      if (respuesta.status === 429) {
        const retryAfterHeader = respuesta.headers.get('Retry-After');
        let retryAfterSeconds = 10; // Valor por defecto si no se especifica Retry-After

        if (retryAfterHeader) {
          retryAfterSeconds = parseInt(retryAfterHeader);
        }

        console.warn(`Solicitud limitada (429). Reintentando después de ${retryAfterSeconds} segundos.`);
        await new Promise(resolve => setTimeout(resolve, retryAfterSeconds * 1000));
        intentos++;
      } else if (!respuesta.ok) {
        throw new Error(`Error al obtener estado de 2FA: ${respuesta.status} ${respuesta.statusText}`);
      } else {
        const datos = await respuesta.json();
        console.log(datos)
        const usuario = datos.value.find(u => u.userPrincipalName === userPrincipalName);

        if (!usuario) {
          return 'Deshabilitado';
        }

        const estado2FA = usuario.authMethods.includes('phone') ? 'Habilitado' : 'Forzado';
        console.log(estado2FA);
        return estado2FA;
      }
    } catch (error) {
      if (intentos >= maxIntentos - 1) {
        throw error;
      } else {
        console.error('Error en solicitud:', error);
        intentos++;
        await new Promise(resolve => setTimeout(resolve, (intentos + 1) * 1000));
      }
    }
  }

  throw new Error(`Se superó el número máximo de intentos (${maxIntentos}) para obtener el estado de 2FA.`);
}


function mostrarDatosEnTabla(datos) {
  $('#tablaDatos').DataTable().destroy();
  $('#tablaDatos tbody').empty();

  const columnas = [
    { data: 'id', title: 'ID' },
    { data: 'displayName', title: 'Nombre Completo' },
    { data: 'mail', title: 'Correo Electrónico' },
    { data: 'jobTitle', title: 'Puesto de Trabajo' },
    { data: 'estado2FA', title: '2FA' } // Nueva columna para el estado de 2FA
  ];

  $('#tablaDatos').DataTable({
    data: datos,
    columns: columnas,
    language: {
      url: '//cdn.datatables.net/plug-ins/1.10.21/i18n/Spanish.json'
    }
  });
}

window.addEventListener('DOMContentLoaded', function () {
  getUsersInDomain();
});

*/