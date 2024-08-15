

  // -------------------  FUNCIONES CRUD - MANIPULACIÓN DE DATOS EN SHAREPOINT -----------------------------//

async function obtenerDatos(url) {
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
        return data.d.results[0];
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




async function actualizar(url, newUrl, empleado, lista) {

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
        const empleadoID = data.d.results[0].ID;
        const spdata = `SP.Data.${lista}ListItem`;
        const dataActualizar = {};

        for (const campo in empleado) {
          dataActualizar[campo] = empleado[campo];
        }
        dataActualizar.__metadata = { type: `${spdata}` };

        const actualizarUrl = `${newUrl}(${empleadoID})`;
        await fetch(actualizarUrl, {
          method: "MERGE",
          headers: {
            Accept: "application/json;odata=verbose",
            "X-RequestDigest": document.getElementById("__REQUESTDIGEST").value,
            "Content-Type": "application/json;odata=verbose",
            "IF-MATCH": "*"
          },
          body: JSON.stringify(dataActualizar)
        });
        return true;
      } else {
        return false;
      }
    } else {
      throw new Error(`Error HTTP: ${response.status}`);
    }
  } catch (error) {
    console.error("Error", error);
    throw error;
  }
}
