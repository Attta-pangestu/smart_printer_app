#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script Pengujian Khusus: Auto Rotation dan Full Page
Untuk menguji metode auto rotation dan full page secara terpisah
pada dokumen Test_print.pdf sebelum implementasi ke aplikasi utama.
"""

import os
import sys
import json
import subprocess
import time
from datetime import datetime

class AutoRotationFullPageTester:
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
                
        print("❌ SumatraPDF tidak ditemukan. Akan menggunakan print default Windows...")
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
    
    def test_auto_rotation_method(self):
        """Metode Auto Rotation - Orientasi optimal untuk penggunaan kertas maksimal"""
        method_name = "AUTO ROTATION METHOD"
        print(f"\n{'='*60}")
        print(f"🔄 PENGUJIAN: {method_name}")
        print(f"{'='*60}")
        
        print("📋 DESKRIPSI METODE:")
        print("   • Dokumen akan dirotasi otomatis ke orientasi terbaik")
        print("   • Menggunakan fit-to-page dengan rotasi landscape")
        print("   • Target: 81.9% penggunaan kertas (berdasarkan analisis)")
        print("   • Ukuran prediksi: 268.8 x 190.0 mm")
        print("   • Tidak ada bagian yang terpotong")
        
        if self.sumatra_path:
            # SumatraPDF dengan fit - akan otomatis rotate untuk orientasi terbaik
            cmd = f'"{self.sumatra_path}" -print-to "{self.printer_name}" -print-settings "fit" -silent "{self.pdf_path}"'
            print("\n🔧 PENGATURAN TEKNIS:")
            print("   • Tool: SumatraPDF")
            print("   • Setting: fit (auto-rotation enabled)")
            print("   • Orientasi: Otomatis (landscape optimal)")
        else:
            # Fallback ke print default Windows
            cmd = f'start /min "" "{self.pdf_path}"'
            print("\n⚠️ FALLBACK: Menggunakan print default Windows")
        
        return self._execute_print_command(method_name, cmd, 
            "Auto rotation untuk penggunaan kertas maksimal tanpa terpotong")
    
    def test_full_page_method(self):
        """Metode Full Page - Mengisi seluruh kertas dengan risiko terpotong"""
        method_name = "FULL PAGE METHOD"
        print(f"\n{'='*60}")
        print(f"📄 PENGUJIAN: {method_name}")
        print(f"{'='*60}")
        
        print("📋 DESKRIPSI METODE:")
        print("   • Dokumen akan di-stretch untuk mengisi seluruh kertas")
        print("   • Menggunakan scaling 137% (berdasarkan analisis)")
        print("   • Target: 100% penggunaan kertas")
        print("   • Ukuran prediksi: 406.0 x 287.0 mm")
        print("   ⚠️ PERINGATAN: 29.7% konten akan terpotong")
        print("   👁️ Konten terlihat: 70.3%")
        
        if self.sumatra_path:
            # Menggunakan scaling tinggi untuk full page
            cmd = f'"{self.sumatra_path}" -print-to "{self.printer_name}" -print-settings "137%" -silent "{self.pdf_path}"'
            print("\n🔧 PENGATURAN TEKNIS:")
            print("   • Tool: SumatraPDF")
            print("   • Setting: 137% scaling")
            print("   • Orientasi: Sesuai dokumen asli")
            print("   ⚠️ RISIKO: Bagian tepi akan terpotong")
        else:
            cmd = f'start /min "" "{self.pdf_path}"'
            print("\n⚠️ FALLBACK: Menggunakan print default Windows")
            print("   (Scaling khusus memerlukan SumatraPDF)")
        
        return self._execute_print_command(method_name, cmd,
            "Full page dengan stretch untuk mengisi seluruh kertas (akan terpotong)")
    
    def _execute_print_command(self, method_name, command, description):
        """Menjalankan perintah cetak dan mencatat hasilnya"""
        print(f"\n💻 COMMAND: {command}")
        print(f"📄 DESKRIPSI: {description}")
        
        # Konfirmasi sebelum mencetak
        print("\n⚠️ KONFIRMASI PENCETAKAN")
        print("📋 Pastikan printer sudah siap dan ada kertas A4")
        response = input(f"\n❓ Lanjutkan pencetakan {method_name}? (y/n): ").lower().strip()
        
        if response != 'y':
            print(f"❌ Pencetakan {method_name} dibatalkan")
            return False
        
        try:
            print(f"\n🚀 Mengirim {method_name} ke printer...")
            
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
                
                # Minta feedback dari user
                print(f"\n📊 EVALUASI HASIL {method_name}:")
                self._get_user_feedback(method_name)
            else:
                print(f"❌ {method_name} gagal: {result.stderr}")
            
            return success
            
        except subprocess.TimeoutExpired:
            print(f"⏰ {method_name} timeout setelah 30 detik")
            return False
        except Exception as e:
            print(f"❌ Error menjalankan {method_name}: {e}")
            return False
    
    def _get_user_feedback(self, method_name):
        """Mengumpulkan feedback dari user tentang hasil cetakan"""
        print("\nSilakan periksa hasil cetakan dan berikan feedback:")
        
        questions = [
            "Apakah dokumen tercetak dengan ukuran yang sesuai? (y/n)",
            "Apakah posisi dokumen sudah terpusat di kertas? (y/n)",
            "Apakah ada bagian dokumen yang terpotong? (y/n)",
            "Apakah kualitas cetakan memuaskan? (y/n)"
        ]
        
        feedback = {}
        for i, question in enumerate(questions, 1):
            response = input(f"{i}. {question}: ").lower().strip()
            feedback[f"question_{i}"] = response == 'y'
        
        # Tambahkan feedback ke hasil
        for result in self.test_results:
            if result['method'] == method_name:
                result['user_feedback'] = feedback
                break
        
        print(f"✅ Feedback untuk {method_name} telah dicatat")
    
    def save_test_results(self):
        """Menyimpan hasil pengujian ke file JSON"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"auto_rotation_fullpage_test_results_{timestamp}.json"
        
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump({
                'pdf_file': self.pdf_path,
                'printer': self.printer_name,
                'sumatra_path': self.sumatra_path,
                'test_timestamp': datetime.now().isoformat(),
                'test_purpose': 'Pengujian terpisah Auto Rotation dan Full Page methods',
                'results': self.test_results
            }, f, indent=2, ensure_ascii=False)
        
        print(f"\n📄 Hasil pengujian disimpan: {filename}")
        return filename
    
    def run_separate_tests(self):
        """Menjalankan pengujian terpisah untuk Auto Rotation dan Full Page"""
        print("="*70)
        print("🧪 PENGUJIAN TERPISAH: AUTO ROTATION & FULL PAGE")
        print("="*70)
        
        if not os.path.exists(self.pdf_path):
            print(f"❌ File PDF tidak ditemukan: {self.pdf_path}")
            return False
        
        if not self.get_default_printer():
            print("❌ Tidak dapat mengakses printer")
            return False
        
        print(f"📁 File PDF: {self.pdf_path}")
        print(f"🖨️ Printer: {self.printer_name}")
        print(f"🔧 SumatraPDF: {'Tersedia' if self.sumatra_path else 'Tidak tersedia'}")
        
        print("\n📋 TUJUAN PENGUJIAN:")
        print("   1. Menguji metode Auto Rotation untuk penggunaan kertas optimal")
        print("   2. Menguji metode Full Page untuk pencetakan maksimal")
        print("   3. Mengumpulkan feedback untuk implementasi ke aplikasi utama")
        
        # Konfirmasi awal
        print("\n⚠️ PERINGATAN PENTING:")
        print("   • Script ini akan mencetak 2 dokumen ke printer")
        print("   • Pastikan printer sudah siap dan ada kertas A4 (minimal 2 lembar)")
        print("   • Anda akan diminta feedback setelah setiap pencetakan")
        
        response = input("\n❓ Lanjutkan pengujian? (y/n): ").lower().strip()
        if response != 'y':
            print("❌ Pengujian dibatalkan")
            return False
        
        # Test 1: Auto Rotation
        print("\n" + "="*50)
        print("📊 TEST 1/2: AUTO ROTATION METHOD")
        print("="*50)
        auto_rotation_success = self.test_auto_rotation_method()
        
        if auto_rotation_success:
            print("\n⏳ Menunggu 10 detik sebelum test berikutnya...")
            time.sleep(10)
        
        # Test 2: Full Page
        print("\n" + "="*50)
        print("📊 TEST 2/2: FULL PAGE METHOD")
        print("="*50)
        full_page_success = self.test_full_page_method()
        
        # Simpan hasil
        self.save_test_results()
        
        # Ringkasan dan rekomendasi
        self.print_final_summary()
        
        return True
    
    def print_final_summary(self):
        """Menampilkan ringkasan final dan rekomendasi"""
        print("\n" + "="*70)
        print("📊 RINGKASAN HASIL PENGUJIAN")
        print("="*70)
        
        successful_tests = [r for r in self.test_results if r['success']]
        
        print(f"✅ Berhasil dicetak: {len(successful_tests)}/2 metode")
        
        for result in self.test_results:
            print(f"\n🔍 {result['method']}:")
            print(f"   Status: {'✅ Berhasil' if result['success'] else '❌ Gagal'}")
            
            if 'user_feedback' in result:
                feedback = result['user_feedback']
                print("   Feedback:")
                print(f"     • Ukuran sesuai: {'✅' if feedback.get('question_1') else '❌'}")
                print(f"     • Posisi terpusat: {'✅' if feedback.get('question_2') else '❌'}")
                print(f"     • Tidak terpotong: {'✅' if not feedback.get('question_3') else '❌'}")
                print(f"     • Kualitas baik: {'✅' if feedback.get('question_4') else '❌'}")
        
        print("\n" + "="*50)
        print("💡 REKOMENDASI UNTUK IMPLEMENTASI")
        print("="*50)
        
        print("\n📋 Berdasarkan hasil pengujian:")
        
        # Analisis hasil untuk rekomendasi
        auto_rotation_feedback = None
        full_page_feedback = None
        
        for result in self.test_results:
            if 'AUTO ROTATION' in result['method'] and 'user_feedback' in result:
                auto_rotation_feedback = result['user_feedback']
            elif 'FULL PAGE' in result['method'] and 'user_feedback' in result:
                full_page_feedback = result['user_feedback']
        
        if auto_rotation_feedback:
            auto_score = sum(auto_rotation_feedback.values()) - auto_rotation_feedback.get('question_3', False)
            print(f"\n🔄 AUTO ROTATION METHOD:")
            print(f"   Score: {auto_score}/4")
            if auto_score >= 3:
                print("   ✅ DIREKOMENDASIKAN untuk implementasi")
                print("   💡 Cocok untuk penggunaan kertas optimal tanpa terpotong")
            else:
                print("   ⚠️ Perlu penyesuaian sebelum implementasi")
        
        if full_page_feedback:
            full_score = sum(full_page_feedback.values()) - full_page_feedback.get('question_3', False)
            print(f"\n📄 FULL PAGE METHOD:")
            print(f"   Score: {full_score}/4")
            if full_score >= 3:
                print("   ✅ DIREKOMENDASIKAN untuk implementasi")
                print("   💡 Cocok untuk pencetakan maksimal (dengan risiko terpotong)")
            else:
                print("   ⚠️ Perlu penyesuaian sebelum implementasi")
        
        print("\n🎯 LANGKAH SELANJUTNYA:")
        print("   1. Review hasil cetakan fisik")
        print("   2. Konfirmasi metode yang akan diimplementasi")
        print("   3. Integrasikan metode pilihan ke aplikasi utama")
        print("   4. Test integrasi dengan aplikasi pencetakan")
        
        print("\n📞 KONFIRMASI IMPLEMENTASI:")
        print("   Silakan konfirmasi metode mana yang akan diimplementasi")
        print("   ke aplikasi pencetakan utama berdasarkan hasil di atas.")

def main():
    pdf_file = r"d:\Gawean Rebinmas\Driver_Epson_L120\test_files\Test_print.pdf"
    
    tester = AutoRotationFullPageTester(pdf_file)
    tester.run_separate_tests()

if __name__ == "__main__":
    main()