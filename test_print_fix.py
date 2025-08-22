
#!/usr/bin/env python3
"""
Script test untuk memverifikasi perbaikan pencetakan
"""

import requests
import json

def test_print_with_new_defaults():
    """Test print dengan pengaturan default yang baru"""
    
    # Pengaturan minimal (menggunakan default yang sudah diperbaiki)
    print_settings = {
        "printer_id": "your_printer_id",
        "copies": 1,
        "paper_size": "A4",
        "orientation": "portrait"
        # fit_to_page, center_horizontally, center_vertically akan menggunakan default True
        # margin akan menggunakan default 0.39 inch (10mm)
    }
    
    print("🧪 Testing dengan pengaturan default yang diperbaiki...")
    print(f"📋 Settings: {json.dumps(print_settings, indent=2)}")
    
    # Kirim request ke server
    # response = requests.post("http://localhost:8000/api/print", json=print_settings)
    
    print("✅ Test selesai. Cek hasil cetak untuk memverifikasi perbaikan.")

if __name__ == "__main__":
    test_print_with_new_defaults()
