// Service Worker لتفعيل خصائص تثبيت التطبيق على الموبايل (PWA)
self.addEventListener('install', function(event) {
  self.skipWaiting();
});

self.addEventListener('activate', function(event) {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', function(event) {
  // تمرير كافة الطلبات مباشرة للشبكة لضمان حداثة البيانات المالية لحظياً
  return;
});
