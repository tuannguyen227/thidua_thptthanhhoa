// static/sw.js
const CACHE_NAME = 'sao-do-offline-v2';

// 1. DANH SÁCH TÀI NGUYÊN "BẢO HIỂM": Phải lưu vào máy để mở App khi không có mạng
const ASSETS_TO_CACHE = [
    '/mobile-sao-do',
    '/static/manifest.json',
    'https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css',
    'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css'
];

// CÀI ĐẶT: Nạp giao diện vào bộ nhớ đệm ngay khi App được cài vào điện thoại
self.addEventListener('install', (event) => {
    self.skipWaiting();
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            console.log('[Service Worker] Đang tải trước giao diện Offline...');
            return cache.addAll(ASSETS_TO_CACHE);
        })
    );
});

// KÍCH HOẠT: Dọn rác các bộ đệm phiên bản cũ
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((cacheNames) => {
            return Promise.all(
                cacheNames.map((cache) => {
                    if (cache !== CACHE_NAME) {
                        console.log('[Service Worker] Đang xóa bộ nhớ đệm cũ...');
                        return caches.delete(cache);
                    }
                })
            );
        }).then(() => self.clients.claim())
    );
});

// KIỂM SOÁT LƯU LƯỢNG MẠNG: Chiến lược "Network First, Fallback to Cache"
self.addEventListener('fetch', (event) => {
    let url = new URL(event.request.url);

    // Bỏ qua các API xử lý dữ liệu động và thao tác POST (sẽ được xử lý ngầm ở IndexedDB)
    if (url.pathname.includes('/export_') || url.pathname.includes('/api/') || event.request.method === 'POST') {
        return; 
    }

    event.respondWith(
        fetch(event.request).catch(() => {
            // Khi điện thoại mất sóng, tự động móc giao diện từ bộ nhớ đệm ra hiển thị
            return caches.match(event.request).then(response => {
                if (response) return response;
                // Nếu rớt mạng và tải trang mới chưa từng lưu, trả về trang /mobile-sao-do mặc định
                return caches.match('/mobile-sao-do'); 
            });
        })
    );
});

// ===================================================================
// BỘ LẮNG NGHE BACKGROUND SYNC (ĐỒNG BỘ NGẦM KHI CÓ MẠNG TRỞ LẠI)
// ===================================================================
self.addEventListener('sync', (event) => {
    if (event.tag === 'sync-sao-do-data') {
        console.log('[Service Worker] 📶 Đã bắt được sóng Wi-Fi/4G! Tiến hành đẩy dữ liệu ngầm lên Server...');
        event.waitUntil(triggerFrontendSync());
    }
});

// Hàm gõ cửa Frontend để ra lệnh xả kho dữ liệu IndexedDB
async function triggerFrontendSync() {
    const clients = await self.clients.matchAll({ includeUncontrolled: true, type: 'window' });
    clients.forEach(client => {
        // Gắn tín hiệu gửi vào giao diện người dùng
        client.postMessage({ type: 'NETWORK_RESTORED_FLUSH_QUEUE' });
    });
}
// ===================================================================
// BỘ LẮNG NGHE THÔNG BÁO ĐẨY (WEB PUSH NOTIFICATIONS)
// ===================================================================
self.addEventListener('push', function(event) {
    if (event.data) {
        const payload = event.data.json();
        const options = {
            body: payload.body,
            icon: '/static/icons/icon-192x192.png', // Thay bằng logo THPT Thanh Hòa
            badge: '/static/icons/badge-icon.png',  // Icon nhỏ trên thanh trạng thái (màu trắng trong suốt)
            vibrate: [200, 100, 200, 100, 200],     // Rung kiểu SOS gây chú ý
            data: { url: '/class-dashboard' }       // Link mở ra khi bấm vào thông báo
        };
        
        event.waitUntil(
            self.registration.showNotification(payload.title, options)
        );
    }
});

// Bắt sự kiện khi người dùng bấm vào thông báo trên màn hình khóa
self.addEventListener('notificationclick', function(event) {
    event.notification.close();
    event.waitUntil(
        clients.openWindow(event.notification.data.url)
    );
});