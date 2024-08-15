


function createList() {
    var url = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists";
    var listData = {
        '__metadata': { 'type': 'SP.List' },
        'AllowContentTypes': true,
        'BaseTemplate': 100,
        'ContentTypesEnabled': true,
        'Description': 'Equipos informaticos entregados a Icontec',
        'Title': 'EquiposInformaticos'
    };

    return fetch(url, {
        method: 'POST',
        headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'X-RequestDigest': document.getElementById('__REQUESTDIGEST').value
        },
        body: JSON.stringify(listData)
    }).then(function(response) {
        if (response.ok) {
            return response.json();
        } else {
            throw new Error('Error al crear la lista: ' + response.statusText);
        }
    });
}


/*

createList().then(function(data) {
    console.log('La lista "' + data.d.Title + '" fue creada correctamente.');
}).catch(function(error) {
    console.log('Error al crear la lista: ' + error.message);
});

*/

function addItemToList() {
    var url = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/getbytitle('EquiposInformaticos34')/items";
    var itemData = {
        '__metadata': { 'type': 'SP.Data.EquiposInformaticos34ListItem' },
        'Title': 'Texto de Prueba'
    };

    return fetch(url, {
        method: 'POST',
        headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'X-RequestDigest': document.getElementById('__REQUESTDIGEST').value
        },
        body: JSON.stringify(itemData)
    }).then(function(response) {
        if (response.ok) {
            return response.json();
        } else {
            throw new Error('Error al agregar elemento a la lista: ' + response.statusText);
        }
    });
}

/*

addItemToList().then(function() {
    console.log('Elemento agregado a la lista.');
}).catch(function(error) {
    console.log('Error al agregar elemento a la lista: ' + error.message);
});

*/