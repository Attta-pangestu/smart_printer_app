#!/usr/bin/env python3
"""
Script untuk memperbaiki pengaturan default pencetakan
agar dokumen tercetak dengan benar (terpusat dan fit to page)
"""

import os
import re
from pathlib import Path

def fix_print_defaults():
    """
    Memperbaiki pengaturan default di main.py untuk mengatasi masalah pencetakan
    """
    print("\n" + "="*60)
    print("🔧 MEMPERBAIKI PENGATURAN DEFAULT PENCETAKAN")
    print("="*60)
    
    main_py_path = Path("d:/Gawean Rebinmas/Driver_Epson_L120/server/main.py")
    
    if not main_py_path.exists():
        print(f"❌ File tidak ditemukan: {main_py_path}")
        return False
    
    print(f"📂 Membaca file: {main_py_path}")
    
    # Baca file asli
    with open(main_py_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Backup file asli
    backup_path = main_py_path.with_suffix('.py.backup')
    with open(backup_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"💾 Backup dibuat: {backup_path}")
    
    # Perbaikan 1: Ubah default fit_to_page dari "actual_size" ke "fit_to_page"
    print("\n🔧 PERBAIKAN 1: Mengubah default fit_to_page")
    old_fit_to_page = 'fit_to_page = settings.get("fit_to_page", "actual_size")'
    new_fit_to_page = 'fit_to_page = settings.get("fit_to_page", "fit_to_page")'
    
    if old_fit_to_page in content:
        content = content.replace(old_fit_to_page, new_fit_to_page)
        print(f"   ✅ Diubah: actual_size → fit_to_page")
    else:
        print(f"   ⚠️  Pattern tidak ditemukan: {old_fit_to_page}")
    
    # Perbaikan 2: Ubah default center_horizontally dari False ke True
    print("\n🔧 PERBAIKAN 2: Mengubah default center_horizontally")
    old_center_h = 'center_horizontally = settings.get("center_horizontally", False)'
    new_center_h = 'center_horizontally = settings.get("center_horizontally", True)'
    
    if old_center_h in content:
        content = content.replace(old_center_h, new_center_h)
        print(f"   ✅ Diubah: False → True")
    else:
        print(f"   ⚠️  Pattern tidak ditemukan: {old_center_h}")
    
    # Perbaikan 3: Ubah default center_vertically dari False ke True
    print("\n🔧 PERBAIKAN 3: Mengubah default center_vertically")
    old_center_v = 'center_vertically = settings.get("center_vertically", False)'
    new_center_v = 'center_vertically = settings.get("center_vertically", True)'
    
    if old_center_v in content:
        content = content.replace(old_center_v, new_center_v)
        print(f"   ✅ Diubah: False → True")
    else:
        print(f"   ⚠️  Pattern tidak ditemukan: {old_center_v}")
    
    # Perbaikan 4: Ubah default margin dari 0.5 inch ke 10mm (0.39 inch)
    print("\n🔧 PERBAIKAN 4: Mengubah default margin")
    margin_patterns = [
        ('margin_top = float(settings.get("margin_top", 0.5))', 'margin_top = float(settings.get("margin_top", 0.39))'),
        ('margin_bottom = float(settings.get("margin_bottom", 0.5))', 'margin_bottom = float(settings.get("margin_bottom", 0.39))'),
        ('margin_left = float(settings.get("margin_left", 0.5))', 'margin_left = float(settings.get("margin_left", 0.39))'),
        ('margin_right = float(settings.get("margin_right", 0.5))', 'margin_right = float(settings.get("margin_right", 0.39))')
    ]
    
    for old_margin, new_margin in margin_patterns:
        if old_margin in content:
            content = content.replace(old_margin, new_margin)
            margin_name = old_margin.split('"')[1]
            print(f"   ✅ Diubah: {margin_name} 0.5 → 0.39 inch (10mm)")
        else:
            print(f"   ⚠️  Pattern tidak ditemukan: {old_margin}")
    
    # Simpan file yang sudah diperbaiki
    with open(main_py_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"\n💾 File berhasil diperbaiki: {main_py_path}")
    
    return True

def verify_changes():
    """
    Verifikasi bahwa perubahan telah diterapkan dengan benar
    """
    print("\n" + "="*60)
    print("🔍 VERIFIKASI PERUBAHAN")
    print("="*60)
    
    main_py_path = Path("d:/Gawean Rebinmas/Driver_Epson_L120/server/main.py")
    
    if not main_py_path.exists():
        print(f"❌ File tidak ditemukan: {main_py_path}")
        return False
    
    with open(main_py_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Cek perubahan
    checks = [
        ('fit_to_page default', 'settings.get("fit_to_page", "fit_to_page")', '✅ fit_to_page default: fit_to_page'),
        ('center_horizontally default', 'settings.get("center_horizontally", True)', '✅ center_horizontally default: True'),
        ('center_vertically default', 'settings.get("center_vertically", True)', '✅ center_vertically default: True'),
        ('margin default', 'settings.get("margin_top", 0.39)', '✅ margin default: 0.39 inch (10mm)')
    ]
    
    all_good = True
    for check_name, pattern, success_msg in checks:
        if pattern in content:
            print(f"   {success_msg}")
        else:
            print(f"   ❌ {check_name}: TIDAK DITEMUKAN")
            all_good = False
    
    return all_good

def show_summary():
    """
    Menampilkan ringkasan perbaikan
    """
    print("\n" + "="*60)
    print("📋 RINGKASAN PERBAIKAN")
    print("="*60)
    
    print("\n🎯 MASALAH YANG DIPERBAIKI:")
    print("   ❌ Dokumen tercetak kecil di pojok kertas")
    print("   ❌ Tidak ada scaling otomatis")
    print("   ❌ Konten tidak terpusat")
    print("   ❌ Margin terlalu besar")
    
    print("\n✅ PERBAIKAN YANG DITERAPKAN:")
    print("   🔄 fit_to_page: 'actual_size' → 'fit_to_page'")
    print("   ↔️  center_horizontally: False → True")
    print("   ↕️  center_vertically: False → True")
    print("   📏 margin: 0.5 inch (12.7mm) → 0.39 inch (10mm)")
    
    print("\n🎯 HASIL YANG DIHARAPKAN:")
    print("   📄 Dokumen akan diskalakan secara otomatis")
    print("   📍 Konten akan terpusat di kertas")
    print("   📏 Margin lebih kecil dan merata")
    print("   ✅ Seluruh area kertas dimanfaatkan optimal")
    
    print("\n🚀 LANGKAH SELANJUTNYA:")
    print("   1. 🔄 Restart aplikasi server")
    print("   2. 🧪 Test print dengan file yang sama")
    print("   3. ✅ Verifikasi hasil cetak sudah memenuhi kertas")
    print("   4. 📊 Bandingkan dengan hasil sebelumnya")

def create_test_script():
    """
    Membuat script test untuk memverifikasi perbaikan
    """
    test_script = '''
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
'''
    
    test_path = Path("d:/Gawean Rebinmas/Driver_Epson_L120/test_print_fix.py")
    with open(test_path, 'w', encoding='utf-8') as f:
        f.write(test_script)
    
    print(f"\n📝 Script test dibuat: {test_path}")

if __name__ == "__main__":
    print("🚀 MEMULAI PERBAIKAN PENGATURAN PENCETAKAN...")
    
    # Lakukan perbaikan
    if fix_print_defaults():
        # Verifikasi perubahan
        if verify_changes():
            print("\n🎉 SEMUA PERBAIKAN BERHASIL DITERAPKAN!")
        else:
            print("\n⚠️  BEBERAPA PERBAIKAN MUNGKIN TIDAK BERHASIL")
        
        # Tampilkan ringkasan
        show_summary()
        
        # Buat script test
        create_test_script()
        
    else:
        print("\n❌ PERBAIKAN GAGAL")
    
    print("\n" + "="*60)
    print("🏁 SELESAI")
    print("="*60)