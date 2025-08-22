#!/usr/bin/env python3
"""
Script untuk menguji berbagai metode pencetakan PDF secara full page
Menguji berbagai pendekatan untuk mencetak dokumen sesuai ukuran kertas penuh
"""

import os
import sys
import json
from datetime import datetime
from pathlib import Path

# Tambahkan path server ke sys.path
server_path = Path(__file__).parent / "server"
sys.path.insert(0, str(server_path))

try:
    import PyPDF2
except ImportError:
    print("PyPDF2 tidak ditemukan. Menggunakan analisis alternatif...")
    PyPDF2 = None

class FullPagePrintTester:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.results = {}
        self.pdf_info = {}
        
    def analyze_pdf(self):
        """Analisis dimensi dan orientasi PDF"""
        print(f"\n=== ANALISIS PDF: {self.pdf_path} ===")
        
        if not os.path.exists(self.pdf_path):
            print(f"❌ File tidak ditemukan: {self.pdf_path}")
            return False
            
        file_size = os.path.getsize(self.pdf_path)
        print(f"📁 Ukuran file: {file_size:,} bytes")
        
        if PyPDF2:
            try:
                with open(self.pdf_path, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    num_pages = len(pdf_reader.pages)
                    
                    if num_pages > 0:
                        page = pdf_reader.pages[0]
                        mediabox = page.mediabox
                        
                        # Konversi dari points ke mm (1 point = 0.352778 mm)
                        width_mm = float(mediabox.width) * 0.352778
                        height_mm = float(mediabox.height) * 0.352778
                        
                        self.pdf_info = {
                            'pages': num_pages,
                            'width_mm': width_mm,
                            'height_mm': height_mm,
                            'width_points': float(mediabox.width),
                            'height_points': float(mediabox.height),
                            'orientation': 'landscape' if width_mm > height_mm else 'portrait'
                        }
                        
                        print(f"📄 Jumlah halaman: {num_pages}")
                        print(f"📏 Dimensi: {width_mm:.1f} x {height_mm:.1f} mm")
                        print(f"📏 Dimensi: {mediabox.width:.1f} x {mediabox.height:.1f} points")
                        print(f"🔄 Orientasi: {self.pdf_info['orientation']}")
                        
                        # Deteksi ukuran kertas standar
                        paper_size = self.detect_paper_size(width_mm, height_mm)
                        print(f"📋 Ukuran kertas: {paper_size}")
                        
                        return True
                        
            except Exception as e:
                print(f"❌ Error membaca PDF: {e}")
                
        # Fallback analysis
        self.pdf_info = {
            'pages': 1,
            'width_mm': 297.0,  # Asumsi A4 landscape
            'height_mm': 210.0,
            'orientation': 'landscape'
        }
        print("⚠️ Menggunakan asumsi A4 landscape")
        return True
        
    def detect_paper_size(self, width_mm, height_mm):
        """Deteksi ukuran kertas berdasarkan dimensi"""
        # Toleransi 5mm untuk variasi
        tolerance = 5
        
        paper_sizes = {
            'A4': [(210, 297), (297, 210)],
            'A3': [(297, 420), (420, 297)],
            'Letter': [(216, 279), (279, 216)],
            'Legal': [(216, 356), (356, 216)]
        }
        
        for size_name, dimensions in paper_sizes.items():
            for w, h in dimensions:
                if (abs(width_mm - w) <= tolerance and abs(height_mm - h) <= tolerance):
                    orientation = 'landscape' if width_mm > height_mm else 'portrait'
                    return f"{size_name} {orientation}"
                    
        return f"Custom ({width_mm:.0f}x{height_mm:.0f}mm)"
        
    def method_1_fit_to_page_center(self):
        """Metode 1: Fit to Page dengan Center aktif"""
        print("\n=== METODE 1: FIT TO PAGE + CENTER ===")
        
        settings = {
            'method_name': 'Fit to Page + Center',
            'fit_to_page': 'fit_to_page',
            'center_horizontally': True,
            'center_vertically': True,
            'margin_top': 0.39,  # 10mm
            'margin_bottom': 0.39,
            'margin_left': 0.39,
            'margin_right': 0.39,
            'custom_scale': 100,
            'description': 'Menyesuaikan dokumen ke ukuran kertas dengan pemusatan otomatis'
        }
        
        # Simulasi perhitungan
        paper_width = 210  # A4 portrait
        paper_height = 297
        margin = 10  # 10mm
        
        printable_width = paper_width - (2 * margin)
        printable_height = paper_height - (2 * margin)
        
        doc_width = self.pdf_info.get('width_mm', 297)
        doc_height = self.pdf_info.get('height_mm', 210)
        
        scale_x = printable_width / doc_width
        scale_y = printable_height / doc_height
        scale = min(scale_x, scale_y)  # Fit to page menggunakan scale terkecil
        
        final_width = doc_width * scale
        final_height = doc_height * scale
        
        # Posisi dengan centering
        pos_x = (paper_width - final_width) / 2
        pos_y = (paper_height - final_height) / 2
        
        settings.update({
            'calculated_scale': f"{scale:.3f} ({scale*100:.1f}%)",
            'final_size': f"{final_width:.1f} x {final_height:.1f} mm",
            'position': f"x={pos_x:.1f}mm, y={pos_y:.1f}mm",
            'paper_usage': f"{(final_width*final_height)/(paper_width*paper_height)*100:.1f}%"
        })
        
        print(f"📐 Scale: {settings['calculated_scale']}")
        print(f"📏 Ukuran final: {settings['final_size']}")
        print(f"📍 Posisi: {settings['position']}")
        print(f"📊 Penggunaan kertas: {settings['paper_usage']}")
        
        self.results['method_1'] = settings
        return settings
        
    def method_2_custom_scaling(self):
        """Metode 2: Custom scaling untuk memaksimalkan ukuran"""
        print("\n=== METODE 2: CUSTOM SCALING MAKSIMAL ===")
        
        # Hitung scale maksimal dengan margin minimal
        paper_width = 210
        paper_height = 297
        min_margin = 5  # 5mm margin minimal
        
        printable_width = paper_width - (2 * min_margin)
        printable_height = paper_height - (2 * min_margin)
        
        doc_width = self.pdf_info.get('width_mm', 297)
        doc_height = self.pdf_info.get('height_mm', 210)
        
        scale_x = printable_width / doc_width
        scale_y = printable_height / doc_height
        max_scale = min(scale_x, scale_y)
        
        settings = {
            'method_name': 'Custom Scaling Maksimal',
            'fit_to_page': 'custom',
            'center_horizontally': True,
            'center_vertically': True,
            'margin_top': 0.197,  # 5mm
            'margin_bottom': 0.197,
            'margin_left': 0.197,
            'margin_right': 0.197,
            'custom_scale': int(max_scale * 100),
            'description': 'Scaling maksimal dengan margin minimal untuk ukuran terbesar'
        }
        
        final_width = doc_width * max_scale
        final_height = doc_height * max_scale
        
        pos_x = (paper_width - final_width) / 2
        pos_y = (paper_height - final_height) / 2
        
        settings.update({
            'calculated_scale': f"{max_scale:.3f} ({max_scale*100:.1f}%)",
            'final_size': f"{final_width:.1f} x {final_height:.1f} mm",
            'position': f"x={pos_x:.1f}mm, y={pos_y:.1f}mm",
            'paper_usage': f"{(final_width*final_height)/(paper_width*paper_height)*100:.1f}%"
        })
        
        print(f"📐 Scale: {settings['calculated_scale']}")
        print(f"📏 Ukuran final: {settings['final_size']}")
        print(f"📍 Posisi: {settings['position']}")
        print(f"📊 Penggunaan kertas: {settings['paper_usage']}")
        
        self.results['method_2'] = settings
        return settings
        
    def method_3_auto_rotation(self):
        """Metode 3: Rotasi otomatis untuk menyesuaikan orientasi"""
        print("\n=== METODE 3: AUTO ROTATION ===")
        
        doc_width = self.pdf_info.get('width_mm', 297)
        doc_height = self.pdf_info.get('height_mm', 210)
        doc_orientation = self.pdf_info.get('orientation', 'landscape')
        
        # Coba kedua orientasi kertas
        paper_portrait = (210, 297)
        paper_landscape = (297, 210)
        
        margin = 10
        
        # Hitung scale untuk portrait
        printable_p = (paper_portrait[0] - 2*margin, paper_portrait[1] - 2*margin)
        scale_p = min(printable_p[0]/doc_width, printable_p[1]/doc_height)
        
        # Hitung scale untuk landscape
        printable_l = (paper_landscape[0] - 2*margin, paper_landscape[1] - 2*margin)
        scale_l = min(printable_l[0]/doc_width, printable_l[1]/doc_height)
        
        # Pilih orientasi yang memberikan scale terbesar
        if scale_p > scale_l:
            best_orientation = 'portrait'
            best_scale = scale_p
            paper_size = paper_portrait
        else:
            best_orientation = 'landscape'
            best_scale = scale_l
            paper_size = paper_landscape
            
        settings = {
            'method_name': 'Auto Rotation',
            'fit_to_page': 'fit_to_page',
            'center_horizontally': True,
            'center_vertically': True,
            'margin_top': 0.39,
            'margin_bottom': 0.39,
            'margin_left': 0.39,
            'margin_right': 0.39,
            'custom_scale': 100,
            'paper_orientation': best_orientation,
            'description': f'Rotasi otomatis ke {best_orientation} untuk scale optimal'
        }
        
        final_width = doc_width * best_scale
        final_height = doc_height * best_scale
        
        pos_x = (paper_size[0] - final_width) / 2
        pos_y = (paper_size[1] - final_height) / 2
        
        settings.update({
            'calculated_scale': f"{best_scale:.3f} ({best_scale*100:.1f}%)",
            'final_size': f"{final_width:.1f} x {final_height:.1f} mm",
            'position': f"x={pos_x:.1f}mm, y={pos_y:.1f}mm",
            'paper_usage': f"{(final_width*final_height)/(paper_size[0]*paper_size[1])*100:.1f}%",
            'recommended_paper': f"{paper_size[0]}x{paper_size[1]}mm {best_orientation}"
        })
        
        print(f"🔄 Orientasi terbaik: {best_orientation}")
        print(f"📐 Scale: {settings['calculated_scale']}")
        print(f"📏 Ukuran final: {settings['final_size']}")
        print(f"📍 Posisi: {settings['position']}")
        print(f"📊 Penggunaan kertas: {settings['paper_usage']}")
        
        self.results['method_3'] = settings
        return settings
        
    def method_4_stretch_fill(self):
        """Metode 4: Stretch/Fill untuk mengisi seluruh kertas"""
        print("\n=== METODE 4: STRETCH TO FILL ===")
        
        paper_width = 210
        paper_height = 297
        margin = 5  # Margin minimal
        
        printable_width = paper_width - (2 * margin)
        printable_height = paper_height - (2 * margin)
        
        doc_width = self.pdf_info.get('width_mm', 297)
        doc_height = self.pdf_info.get('height_mm', 210)
        
        # Stretch untuk mengisi area cetak (mungkin mengubah aspect ratio)
        scale_x = printable_width / doc_width
        scale_y = printable_height / doc_height
        
        # Gunakan scale yang lebih besar untuk fill (bukan fit)
        fill_scale = max(scale_x, scale_y)
        
        settings = {
            'method_name': 'Stretch to Fill',
            'fit_to_page': 'fill',  # atau 'custom' dengan scale tinggi
            'center_horizontally': True,
            'center_vertically': True,
            'margin_top': 0.197,
            'margin_bottom': 0.197,
            'margin_left': 0.197,
            'margin_right': 0.197,
            'custom_scale': int(fill_scale * 100),
            'description': 'Mengisi seluruh area cetak, mungkin memotong sebagian konten'
        }
        
        final_width = doc_width * fill_scale
        final_height = doc_height * fill_scale
        
        # Hitung area yang terpotong
        overflow_x = max(0, final_width - printable_width)
        overflow_y = max(0, final_height - printable_height)
        
        pos_x = (paper_width - final_width) / 2
        pos_y = (paper_height - final_height) / 2
        
        settings.update({
            'calculated_scale': f"{fill_scale:.3f} ({fill_scale*100:.1f}%)",
            'final_size': f"{final_width:.1f} x {final_height:.1f} mm",
            'position': f"x={pos_x:.1f}mm, y={pos_y:.1f}mm",
            'paper_usage': f"{min(100, (final_width*final_height)/(paper_width*paper_height)*100):.1f}%",
            'overflow': f"x={overflow_x:.1f}mm, y={overflow_y:.1f}mm",
            'content_visible': f"{max(0, 100 - (overflow_x+overflow_y)/(final_width+final_height)*100):.1f}%"
        })
        
        print(f"📐 Scale: {settings['calculated_scale']}")
        print(f"📏 Ukuran final: {settings['final_size']}")
        print(f"📍 Posisi: {settings['position']}")
        print(f"📊 Penggunaan kertas: {settings['paper_usage']}")
        if overflow_x > 0 or overflow_y > 0:
            print(f"⚠️ Area terpotong: {settings['overflow']}")
            print(f"👁️ Konten terlihat: {settings['content_visible']}")
        
        self.results['method_4'] = settings
        return settings
        
    def generate_test_configs(self):
        """Generate konfigurasi untuk testing"""
        print("\n=== GENERATE TEST CONFIGURATIONS ===")
        
        configs = {}
        for method_id, method_data in self.results.items():
            config = {
                'pdf_path': self.pdf_path,
                'printer_name': 'EPSON L120 Series',  # Sesuaikan dengan printer
                'settings': {
                    'fit_to_page': method_data.get('fit_to_page', 'fit_to_page'),
                    'center_horizontally': method_data.get('center_horizontally', True),
                    'center_vertically': method_data.get('center_vertically', True),
                    'margin_top': method_data.get('margin_top', 0.39),
                    'margin_bottom': method_data.get('margin_bottom', 0.39),
                    'margin_left': method_data.get('margin_left', 0.39),
                    'margin_right': method_data.get('margin_right', 0.39),
                    'custom_scale': method_data.get('custom_scale', 100)
                },
                'expected_result': {
                    'scale': method_data.get('calculated_scale', 'N/A'),
                    'final_size': method_data.get('final_size', 'N/A'),
                    'paper_usage': method_data.get('paper_usage', 'N/A')
                }
            }
            configs[method_id] = config
            
        return configs
        
    def save_results(self):
        """Simpan hasil analisis dan konfigurasi"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        report = {
            'timestamp': timestamp,
            'pdf_file': self.pdf_path,
            'pdf_info': self.pdf_info,
            'methods': self.results,
            'test_configs': self.generate_test_configs(),
            'summary': self.generate_summary()
        }
        
        report_file = f"full_page_print_report_{timestamp}.json"
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
            
        print(f"\n📄 Laporan disimpan: {report_file}")
        return report_file
        
    def generate_summary(self):
        """Generate ringkasan rekomendasi"""
        summary = {
            'best_method': None,
            'best_paper_usage': 0,
            'recommendations': []
        }
        
        best_usage = 0
        best_method = None
        
        for method_id, method_data in self.results.items():
            usage_str = method_data.get('paper_usage', '0%')
            usage = float(usage_str.replace('%', ''))
            
            if usage > best_usage:
                best_usage = usage
                best_method = method_id
                
        summary['best_method'] = best_method
        summary['best_paper_usage'] = best_usage
        
        # Rekomendasi berdasarkan analisis
        if self.pdf_info.get('orientation') == 'landscape':
            summary['recommendations'].append(
                "Dokumen landscape: Pertimbangkan Method 3 (Auto Rotation) untuk hasil optimal"
            )
            
        if best_usage < 50:
            summary['recommendations'].append(
                "Penggunaan kertas rendah: Coba Method 2 (Custom Scaling) atau Method 4 (Stretch Fill)"
            )
            
        summary['recommendations'].append(
            "Method 1 (Fit to Page) paling aman untuk mempertahankan proporsi dokumen"
        )
        
        return summary
        
    def run_all_tests(self):
        """Jalankan semua test metode"""
        print("🚀 MEMULAI PENGUJIAN METODE PENCETAKAN FULL PAGE")
        print("=" * 60)
        
        # Analisis PDF
        if not self.analyze_pdf():
            return False
            
        # Jalankan semua metode
        self.method_1_fit_to_page_center()
        self.method_2_custom_scaling()
        self.method_3_auto_rotation()
        self.method_4_stretch_fill()
        
        # Simpan hasil
        report_file = self.save_results()
        
        # Tampilkan ringkasan
        self.display_summary()
        
        return True
        
    def display_summary(self):
        """Tampilkan ringkasan hasil"""
        print("\n" + "=" * 60)
        print("📊 RINGKASAN HASIL PENGUJIAN")
        print("=" * 60)
        
        summary = self.generate_summary()
        
        print(f"\n🏆 Metode terbaik: {summary['best_method']}")
        print(f"📊 Penggunaan kertas terbaik: {summary['best_paper_usage']:.1f}%")
        
        print("\n💡 REKOMENDASI:")
        for i, rec in enumerate(summary['recommendations'], 1):
            print(f"   {i}. {rec}")
            
        print("\n📋 PERBANDINGAN METODE:")
        for method_id, method_data in self.results.items():
            print(f"\n   {method_data['method_name']}:")
            print(f"   - Scale: {method_data.get('calculated_scale', 'N/A')}")
            print(f"   - Ukuran: {method_data.get('final_size', 'N/A')}")
            print(f"   - Penggunaan kertas: {method_data.get('paper_usage', 'N/A')}")
            
        print("\n" + "=" * 60)
        print("✅ PENGUJIAN SELESAI")
        print("\n📝 Langkah selanjutnya:")
        print("   1. Review konfigurasi yang dihasilkan")
        print("   2. Test cetak dengan metode pilihan")
        print("   3. Verifikasi hasil cetakan")
        print("   4. Sesuaikan pengaturan jika diperlukan")

def main():
    pdf_path = r"d:\Gawean Rebinmas\Driver_Epson_L120\test_files\Test_print.pdf"
    
    tester = FullPagePrintTester(pdf_path)
    success = tester.run_all_tests()
    
    if success:
        print("\n🎉 Pengujian berhasil diselesaikan!")
    else:
        print("\n❌ Pengujian gagal!")
        
if __name__ == "__main__":
    main()