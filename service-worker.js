self.addEventListener('install', function(e) {
  console.log('Service Worker: Installed');
});

self.addEventListener('activate', function(e) {
  console.log('Service Worker: Activated');
});

self.addEventListener('fetch', function(e) {
  // Aquí puedes gestionar la cache y recursos offline
});
