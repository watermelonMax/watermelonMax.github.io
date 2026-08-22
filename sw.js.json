const CACHE_NAME = "english-word-v1";

const STATIC_FILES = [
    "/",
    "/data/words.xlsx",
    "/data/wrong_words.xlsx",
    "/icons/icon-192.png",
    "/icons/icon-512.png"
];


self.addEventListener("install", event => {

    event.waitUntil(

        caches
        .open(CACHE_NAME)
        .then(cache => {

            return cache.addAll(STATIC_FILES);

        })

    );

});


self.addEventListener("fetch", event => {

    event.respondWith(

        caches
        .match(event.request)
        .then(cached => {

            return cached || fetch(event.request);

        })

    );

});