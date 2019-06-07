const CACHE_NAME = 'Clare',
urlsToCache = [
    './',
    'https://fonts.googleapis.com/icon?family=Material+Icons',
    './css/materialize.min.css',
    './stylesheet.css',
    './script.js',
    './images/favicon.png',
    './manifest.json',
    './res/Wall1.png',
    './res/Wall2.png',
    './res/Wall3.png'
]

// Instala y almacena en cache los archivos estáticos de arriba
self.addEventListener('install', e => {

    // Espera hasta que el evento abra el cache
    e.waitUntil(
        caches.open(CACHE_NAME) // Y abre la version de cache
        .then(cache => { // Administramos la entada de las url
            return cache.addAll(urlsToCache)
            .then(() => self.skipWaiting()) //Que espere
            .catch(err => console.log('Error al cargar: ', err)) //Si no hay url específica
        })
    )
})

// Este nos ayuda cuando perdemos la conexion a internet, busca los recursos en cache y los carga
self.addEventListener('activate', e => {

    // Nos avisara si los archivos en cache se repiten
    const cacheWhiteList = [CACHE_NAME]

    //Espera y ve las llaves del cache que han sufrido modificaciones
    e.waitUntil(
        caches.keys()
        .then(cachesNames => {
            cachesNames.map(cacheName => { //Evalua el cache y lo que haya cambiado lo elimina
                if (cacheWhiteList.indexOf(cacheName) === -1) {
                    return caches.delete(cacheName)
                }
            })
        })
        // Indica que debe activar el cache actual
        .then(() => self.clients.claim())
    )
})

// Recupera los archivos cuando regrese la conexión, encuentra una url real
self.addEventListener('fetch', e => {

    // Mira si recuperó el recurso del cache
    e.respondWith(
        caches.match(e.request)
        .then( res => {
            if (res) {
                return res
            }

            // Recupera la petición de url
            return fetch(e.request)
        })
    )
})
