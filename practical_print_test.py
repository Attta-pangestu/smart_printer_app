#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script Pengujian Praktis untuk Mencetak PDF Full Page
Berdasarkan analisis dari full_page_print_methods.py

Script ini akan melakukan pencetakan aktual menggunakan berbagai metode
yang telah dianalisis sebelumnya.
"""

import os
import sys
import json
import subprocess
import time
from datetime import datetime

class PracticalPrintTester:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.printer_name = None
        self.sumatra_path = self.find_sumatra_pdf()
        self.test_results = []
        
    def find_sumatra_pdf(self):
        """Mencari SumatraPDF di berbagai lokasi umum"""
        possible_paths = [
            r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
            r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
            r"D:\Program Files\SumatraPDF\SumatraPDF.exe",
            os.path.join(os.path.dirname(__file__), 'print_tools', 'SumatraPDF-3.4.6-64.exe'),
            "SumatraPDF.exe"  # Jika ada di PATH
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                print(f"✅ SumatraPDF ditemukan: {path}")
                return path
                
        print("❌ SumatraPDF tidak ditemukan. Mencoba menggunakan print default Windows...")
        return None
    
    def get_default_printer(self):
        """Mendapatkan printer default sistem"""
        try:
            import win32print
            self.printer_name = win32print.GetDefaultPrinter()
            print(f"🖨️ Printer default: {self.printer_name}")
            return True
        except ImportError:
            print("❌ Modul win32print tidak tersedia. Menggunakan printer default sistem.")
            self.printer_name = "default"
            return True
        except Exception as e:
            print(f"❌ Error mendapatkan printer default: {e}")
            return False
    
    def print_method_1_fit_to_page(self):
        """Metode 1: Fit to Page + Center (Paling Aman)"""
        method_name = "Method 1: Fit to Page + Center"
        print(f"\n🔄 Menjalankan {method_name}...")
        
        if self.sumatra_path:
            # Menggunakan SumatraPDF dengan opsi fit
            cmd = f'"{self.sumatra_path}" -print-to "{self.printer_name}" -print-settings "fit" -silent "{self.pdf_path}"'
        else:
            # Fallback ke print default Windows
            cmd = f'start /min "" "{self.pdf_path}"'
        
        return self._execute_print_command(method_name, cmd, 
            "Mencetak dengan fit to page untuk mempertahankan proporsi dokumen")
    
    def print_method_2_custom_scaling(self):
        """Metode 2: Custom Scaling (67.3% berdasarkan analisis)"""
        method_name = "Method 2: Custom Scaling (67%)"
        print(f"\n🔄 Menjalankan {method_name}...")
        
        if self.sumatra_path:
            # Menggunakan scaling 67% berdasarkan analisis
            cmd = f'"{self.sumatra_path}" -print-to "{self.printer_name}" -print-settings "67%" -silent "{self.pdf_path}"'
        else:
            # Fallback - tidak bisa set custom scaling tanpa SumatraPDF
            cmd = f'start /min "" "{self.pdf_path}"'
            print("⚠️ Custom scaling memerlukan SumatraPDF. Menggunakan default.")
        
        return self._execute_print_command(method_name, cmd,
            "Mencetak dengan scaling 67% untuk ukuran lebih besar")
    
    def print_method_3_auto_rotation(self):
        """Metode 3: Auto Rotation (90.5% dengan rotasi landscape)"""
        method_name = "Method 3: Auto Rotation (Landscape)"
        print(f"\n🔄 Menjalankan {method_name}...")
        
        if self.sumatra_path:
            # SumatraPDF biasanya auto-rotate, tapi kita bisa coba dengan fit
            cmd = f'"{self.sumatra_path}" -print-to "{self.printer_name}" -print-settings "fit" -silent "{self.pdf_path}"'
            print("📝 Catatan: SumatraPDF akan otomatis menyesuaikan orientasi")
        else:
            cmd = f'start /min "" "{self.pdf_path}"'
        
        return self._execute_print_command(method_name, cmd,
            "Mencetak dengan rotasi otomatis untuk penggunaan kertas maksimal")
    
    def print_method_4_stretch_fill(self):
        """Metode 4: Stretch to Fill (136.6% - akan terpotong)"""
        method_name = "Method 4: Stretch to Fill (Terpotong)"
        print(f"\n🔄 Menjalankan {method_name}...")
        
        if self.sumatra_path:
            # Menggunakan scaling tinggi - akan terpotong
            cmd = f'"{self.sumatra_path}" -print-to "{self.printer_name}" -print-settings "137%" -silent "{self.pdf_path}"'
            print("⚠️ PERINGATAN: Metode ini akan memotong sebagian dokumen!")
        else:
            cmd = f'start /min "" "{self.pdf_path}"'
            print("⚠️ Stretch to fill memerlukan SumatraPDF. Menggunakan default.")
        
        return self._execute_print_command(method_name, cmd,
            "Mencetak dengan stretch untuk mengisi seluruh kertas (akan terpotong)")
    
    def _execute_print_command(self, method_name, command, description):
        """Menjalankan perintah cetak dan mencatat hasilnya"""
        print(f"📄 {description}")
        print(f"💻 Command: {command}")
        
        try:
            # Jalankan perintah
            result = subprocess.run(command, shell=True, capture_output=True, text=True, timeout=30)
            
            success = result.returncode == 0
            
            test_result = {
                'method': method_name,
                'description': description,
                'command': command,
                'success': success,
                'timestamp': datetime.now().isoformat(),
                'return_code': result.returncode,
                'stdout': result.stdout,
                'stderr': result.stderr
            }
            
            self.test_results.append(test_result)
            
            if success:
                print(f"✅ {method_name} berhasil dikirim ke printer")
                print("📋 Silakan periksa hasil cetakan fisik")
            else:
                print(f"❌ {method_name} gagal: {result.stderr}")
            
            return success
            
        except subprocess.TimeoutExpired:
            print(f"⏰ {method_name} timeout setelah 30 detik")
            return False
        except Exception as e:
            print(f"❌ Error menjalankan {method_name}: {e}")
            return False
    
    def save_test_results(self):
        """Menyimpan hasil pengujian ke file JSON"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"practical_print_results_{timestamp}.json"
        
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump({
                'pdf_file': self.pdf_path,
                'printer': self.printer_name,
                'sumatra_path': self.sumatra_path,
                'test_timestamp': datetime.now().isoformat(),
                'results': self.test_results
            }, f, indent=2, ensure_ascii=False)
        
        print(f"\n📄 Hasil pengujian disimpan: {filename}")
        return filename
    
    def run_all_tests(self):
        """Menjalankan semua metode pengujian"""
        print("="*60)
        print("🚀 MEMULAI PENGUJIAN PRAKTIS PENCETAKAN FULL PAGE")
        print("="*60)
        
        if not os.path.exists(self.pdf_path):
            print(f"❌ File PDF tidak ditemukan: {self.pdf_path}")
            return False
        
        if not self.get_default_printer():
            print("❌ Tidak dapat mengakses printer")
            return False
        
        print(f"📁 File PDF: {self.pdf_path}")
        print(f"🖨️ Printer: {self.printer_name}")
        print(f"🔧 SumatraPDF: {'Tersedia' if self.sumatra_path else 'Tidak tersedia'}")
        
        # Konfirmasi dari user
        print("\n⚠️ PERINGATAN: Script ini akan mengirim dokumen ke printer!")
        print("📋 Pastikan printer sudah siap dan ada kertas yang cukup.")
        
        response = input("\n❓ Lanjutkan pengujian? (y/n): ").lower().strip()
        if response != 'y':
            print("❌ Pengujian dibatalkan oleh user")
            return False
        
        # Jalankan semua metode dengan jeda
        methods = [
            self.print_method_1_fit_to_page,
            self.print_method_2_custom_scaling,
            self.print_method_3_auto_rotation,
            self.print_method_4_stretch_fill
        ]
        
        for i, method in enumerate(methods, 1):
            print(f"\n{'='*50}")
            print(f"📊 PENGUJIAN {i}/{len(methods)}")
            print(f"{'='*50}")
            
            method()
            
            # Jeda antar metode (kecuali yang terakhir)
            if i < len(methods):
                print("\n⏳ Menunggu 5 detik sebelum metode berikutnya...")
                time.sleep(5)
        
        # Simpan hasil
        self.save_test_results()
        
        # Ringkasan
        self.print_summary()
        
        return True
    
    def print_summary(self):
        """Menampilkan ringkasan hasil pengujian"""
        print("\n" + "="*60)
        print("📊 RINGKASAN HASIL PENGUJIAN PRAKTIS")
        print("="*60)
        
        successful_tests = [r for r in self.test_results if r['success']]
        failed_tests = [r for r in self.test_results if not r['success']]
        
        print(f"✅ Berhasil: {len(successful_tests)}/{len(self.test_results)} metode")
        print(f"❌ Gagal: {len(failed_tests)}/{len(self.test_results)} metode")
        
        if successful_tests:
            print("\n🎉 Metode yang berhasil:")
            for result in successful_tests:
                print(f"   ✓ {result['method']}")
        
        if failed_tests:
            print("\n⚠️ Metode yang gagal:")
            for result in failed_tests:
                print(f"   ✗ {result['method']}")
        
        print("\n📋 LANGKAH SELANJUTNYA:")
        print("   1. Periksa hasil cetakan fisik dari setiap metode")
        print("   2. Bandingkan kualitas dan ukuran cetakan")
        print("   3. Pilih metode terbaik berdasarkan kebutuhan")
        print("   4. Dokumentasikan metode pilihan untuk penggunaan selanjutnya")
        
        print("\n💡 REKOMENDASI BERDASARKAN ANALISIS SEBELUMNYA:")
        print("   • Method 1 (Fit to Page): Paling aman, mempertahankan proporsi")
        print("   • Method 3 (Auto Rotation): Penggunaan kertas terbaik (81.9%)")
        print("   • Method 4 (Stretch): Full page tapi akan terpotong (70.3% terlihat)")

def main():
    pdf_file = r"d:\Gawean Rebinmas\Driver_Epson_L120\test_files\Test_print.pdf"
    
    tester = PracticalPrintTester(pdf_file)
    tester.run_all_tests()

if __name__ == "__main__":
    main()