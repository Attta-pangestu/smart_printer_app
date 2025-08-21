# Printer Sharing System

Sistem client-server untuk berbagi printer di jaringan LAN yang memungkinkan kontrol penuh terhadap pengaturan printing.

## Fitur

- **Network Discovery**: Otomatis menemukan printer di jaringan LAN
- **Web Interface**: Interface web yang mudah digunakan untuk mengatur printing
- **Print Control**: Kontrol lengkap untuk:
  - Mode warna (Color/Grayscale/Black & White)
  - Jumlah copy
  - Layout dan orientasi
  - Ukuran kertas
  - Kualitas print
  - Duplex printing
- **Client Driver**: Driver client untuk komunikasi dengan server printer
- **Job Management**: Tracking dan history pekerjaan printing

## Arsitektur

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   Client App    │───▶│  Server API     │───▶│   Printer       │
│  (Web/Desktop)  │    │  (FastAPI)      │    │  (Network)      │
└─────────────────┘    └─────────────────┘    └─────────────────┘
```

## Struktur Proyek

```
Driver_Epson_L120/
├── server/                 # Backend server
│   ├── api/               # API endpoints
│   ├── models/            # Data models
│   ├── services/          # Business logic
│   └── main.py           # Server entry point
├── client/                # Client driver
│   ├── discovery/        # Network discovery
│   ├── printer/          # Printer communication
│   └── main.py          # Client entry point
├── web/                   # Web interface
│   ├── static/           # CSS, JS, images
│   ├── templates/        # HTML templates
│   └── app.py           # Web app
├── config.yaml           # Configuration
├── requirements.txt      # Dependencies
└── README.md            # Documentation
```

## Instalasi

1. Clone repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Konfigurasi `config.yaml` sesuai kebutuhan
4. Jalankan server:
   ```bash
   python server/main.py
   ```
5. Akses web interface di `http://localhost:8080`

## Penggunaan

### Server (Komputer dengan Printer)
1. Jalankan server printer sharing
2. Server akan otomatis mendeteksi printer yang terhubung
3. Web interface tersedia untuk monitoring

### Client (Komputer lain di LAN)
1. Jalankan client discovery untuk menemukan server
2. Gunakan web interface atau API untuk printing
3. Atur konfigurasi printing sesuai kebutuhan

## API Endpoints

- `GET /api/printers` - List printer yang tersedia
- `POST /api/print` - Submit job printing
- `GET /api/jobs` - List job printing
- `GET /api/status` - Status server dan printer

## Konfigurasi Printing

- **Color Mode**: Color, Grayscale, Black & White
- **Copies**: Jumlah copy (1-999)
- **Paper Size**: A4, A3, Letter, Legal, dll
- **Orientation**: Portrait, Landscape
- **Quality**: Draft, Normal, High
- **Duplex**: Single-sided, Double-sided

## Lisensi

MIT License