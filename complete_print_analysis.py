import fitz  # PyMuPDF
import os

def analyze_print_problem():
    """
    Analisis lengkap masalah pencetakan PDF
    """
    print("\n" + "="*60)
    print("🔍 ANALISIS MASALAH PENCETAKAN PDF")
    print("="*60)
    
    # Coba beberapa lokasi file PDF
    possible_paths = [
        "C:\\Users\\nbgmf\\Downloads\\Test_print.pdf",
        "d:\\Gawean Rebinmas\\Driver_Epson_L120\\temp\\Test_print.pdf",
        "d:\\Gawean Rebinmas\\Driver_Epson_L120\\Test_print.pdf"
    ]
    
    pdf_path = None
    for path in possible_paths:
        if os.path.exists(path):
            pdf_path = path
            print(f"✅ File ditemukan: {path}")
            break
    
    if not pdf_path:
        print("❌ File Test_print.pdf tidak ditemukan di lokasi manapun")
        print("📍 Lokasi yang dicoba:")
        for path in possible_paths:
            print(f"   - {path}")
        print("\n💡 Silakan pastikan file ada di salah satu lokasi tersebut")
        return analyze_without_file()
    
    try:
        # Analisis file PDF
        doc = fitz.open(pdf_path)
        page = doc[0]
        
        print(f"\n📄 INFORMASI DOKUMEN ASLI:")
        print(f"   📁 File: {os.path.basename(pdf_path)}")
        print(f"   📐 Dimensi: {page.rect.width:.1f} x {page.rect.height:.1f} pt")
        print(f"   📏 Ukuran: {page.rect.width * 0.352778:.1f} x {page.rect.height * 0.352778:.1f} mm")
        
        # Tentukan orientasi
        if page.rect.width > page.rect.height:
            orientation = "Landscape"
            print(f"   🔄 Orientasi: {orientation} (Lebar > Tinggi)")
        else:
            orientation = "Portrait"
            print(f"   🔄 Orientasi: {orientation} (Tinggi > Lebar)")
        
        # Analisis konten
        images = page.get_images()
        text_blocks = page.get_text("blocks")
        
        print(f"\n🖼️  ANALISIS KONTEN:")
        print(f"   📷 Jumlah gambar: {len(images)}")
        print(f"   📝 Jumlah blok teks: {len(text_blocks)}")
        
        if images:
            for img_index, img in enumerate(images[:3]):  # Maksimal 3 gambar
                try:
                    xref = img[0]
                    pix = fitz.Pixmap(doc, xref)
                    print(f"   🖼️  Gambar {img_index + 1}: {pix.width} x {pix.height} piksel")
                    print(f"      Resolusi: {pix.width * pix.height:,} piksel total")
                    pix = None
                except:
                    print(f"   🖼️  Gambar {img_index + 1}: Tidak dapat dianalisis")
        
        doc.close()
        
        # Lakukan simulasi pencetakan
        simulate_printing_scenarios(page.rect.width, page.rect.height, orientation)
        
    except Exception as e:
        print(f"❌ Error menganalisis PDF: {str(e)}")
        return analyze_without_file()

def analyze_without_file():
    """
    Analisis berdasarkan informasi umum tanpa file
    """
    print(f"\n📄 ANALISIS BERDASARKAN INFORMASI UMUM:")
    print(f"   📐 Asumsi: Dokumen A4 Landscape (297 x 210 mm)")
    print(f"   🎯 Target: Pencetakan A4 Portrait (210 x 297 mm)")
    
    # Simulasi dengan dimensi standar A4 landscape
    simulate_printing_scenarios(842.40, 595.44, "Landscape")

def simulate_printing_scenarios(doc_width_pt, doc_height_pt, orientation):
    """
    Simulasi berbagai skenario pencetakan
    """
    print(f"\n" + "="*60)
    print(f"📋 SIMULASI SKENARIO PENCETAKAN")
    print(f"="*60)
    
    # Ukuran kertas A4 portrait (target pencetakan)
    a4_width_pt = 595.28  # 210mm
    a4_height_pt = 841.89  # 297mm
    
    print(f"\n🎯 TARGET KERTAS (A4 Portrait):")
    print(f"   📐 Dimensi: {a4_width_pt:.1f} x {a4_height_pt:.1f} pt")
    print(f"   📏 Ukuran: 210 x 297 mm")
    
    print(f"\n📄 DOKUMEN ASLI ({orientation}):")
    print(f"   📐 Dimensi: {doc_width_pt:.1f} x {doc_height_pt:.1f} pt")
    print(f"   📏 Ukuran: {doc_width_pt * 0.352778:.1f} x {doc_height_pt * 0.352778:.1f} mm")
    
    # Skenario 1: Tanpa scaling (masalah utama)
    print(f"\n❌ SKENARIO 1 - TANPA SCALING (MASALAH SAAT INI):")
    print(f"   📍 Dokumen ditempatkan langsung di pojok kiri atas")
    
    if doc_width_pt > a4_width_pt:
        overflow_width = doc_width_pt - a4_width_pt
        print(f"   ⚠️  Lebar dokumen ({doc_width_pt:.0f}pt) > Lebar kertas ({a4_width_pt:.0f}pt)")
        print(f"   ✂️  Terpotong: {overflow_width:.0f}pt ({overflow_width * 0.352778:.1f}mm) di kanan")
    
    if doc_height_pt > a4_height_pt:
        overflow_height = doc_height_pt - a4_height_pt
        print(f"   ⚠️  Tinggi dokumen ({doc_height_pt:.0f}pt) > Tinggi kertas ({a4_height_pt:.0f}pt)")
        print(f"   ✂️  Terpotong: {overflow_height:.0f}pt ({overflow_height * 0.352778:.1f}mm) di bawah")
    
    visible_width = min(doc_width_pt, a4_width_pt)
    visible_height = min(doc_height_pt, a4_height_pt)
    visible_ratio = (visible_width * visible_height) / (doc_width_pt * doc_height_pt)
    
    print(f"   👁️  Area yang terlihat: {visible_ratio*100:.1f}% dari dokumen asli")
    print(f"   📍 Posisi: Pojok kiri atas kertas")
    
    # Skenario 2: Dengan fit to page
    margin_pt = 20 * 2.834645  # 20mm to points
    available_width = a4_width_pt - (2 * margin_pt)
    available_height = a4_height_pt - (2 * margin_pt)
    
    scale_x = available_width / doc_width_pt
    scale_y = available_height / doc_height_pt
    scale_factor = min(scale_x, scale_y)
    
    scaled_width = doc_width_pt * scale_factor
    scaled_height = doc_height_pt * scale_factor
    
    print(f"\n✅ SKENARIO 2 - DENGAN FIT TO PAGE (SOLUSI):")
    print(f"   📏 Margin yang diinginkan: 20mm")
    print(f"   📐 Area tersedia: {available_width:.1f} x {available_height:.1f} pt")
    print(f"   🔢 Scale X: {scale_x:.3f} ({scale_x*100:.1f}%)")
    print(f"   🔢 Scale Y: {scale_y:.3f} ({scale_y*100:.1f}%)")
    print(f"   ⚖️  Scale factor: {scale_factor:.3f} ({scale_factor*100:.1f}%)")
    print(f"   📐 Ukuran setelah scaling: {scaled_width:.1f} x {scaled_height:.1f} pt")
    print(f"   📏 Ukuran dalam mm: {scaled_width*0.352778:.1f} x {scaled_height*0.352778:.1f} mm")
    
    # Hitung posisi dengan centering
    pos_x = (a4_width_pt - scaled_width) / 2
    pos_y = (a4_height_pt - scaled_height) / 2
    
    print(f"\n📍 POSISI DENGAN CENTERING:")
    print(f"   ↔️  X: {pos_x:.1f} pt ({pos_x*0.352778:.1f} mm dari kiri)")
    print(f"   ↕️  Y: {pos_y:.1f} pt ({pos_y*0.352778:.1f} mm dari atas)")
    
    # Margin aktual
    margin_left = pos_x
    margin_top = pos_y
    margin_right = a4_width_pt - (pos_x + scaled_width)
    margin_bottom = a4_height_pt - (pos_y + scaled_height)
    
    print(f"\n📏 MARGIN AKTUAL SETELAH PERBAIKAN:")
    print(f"   ⬅️  Kiri: {margin_left:.1f} pt ({margin_left*0.352778:.1f} mm)")
    print(f"   ⬆️  Atas: {margin_top:.1f} pt ({margin_top*0.352778:.1f} mm)")
    print(f"   ➡️  Kanan: {margin_right:.1f} pt ({margin_right*0.352778:.1f} mm)")
    print(f"   ⬇️  Bawah: {margin_bottom:.1f} pt ({margin_bottom*0.352778:.1f} mm)")
    
    # Perbandingan
    improvement = (scale_factor * 100) / (visible_ratio * 100) if visible_ratio > 0 else float('inf')
    print(f"\n📊 PERBANDINGAN:")
    print(f"   ❌ Tanpa scaling: {visible_ratio*100:.1f}% dokumen terlihat")
    print(f"   ✅ Dengan scaling: {scale_factor*100:.1f}% dokumen terlihat (100% konten)")
    if improvement != float('inf'):
        print(f"   📈 Peningkatan: {improvement:.1f}x lebih baik")

def provide_solutions():
    """
    Memberikan solusi untuk masalah pencetakan
    """
    print(f"\n" + "="*60)
    print(f"🔧 SOLUSI MASALAH PENCETAKAN")
    print(f"="*60)
    
    print(f"\n🎯 PENYEBAB UTAMA MASALAH:")
    print(f"   1. 📄 Dokumen dalam format Landscape (lebar > tinggi)")
    print(f"   2. 🖨️  Printer diatur untuk Portrait (tinggi > lebar)")
    print(f"   3. ❌ Tidak ada scaling otomatis")
    print(f"   4. 📍 Konten ditempatkan di pojok kiri atas")
    print(f"   5. ✂️  Sebagian besar konten terpotong")
    
    print(f"\n💡 SOLUSI YANG HARUS DITERAPKAN:")
    print(f"   ✅ 1. AKTIFKAN 'Fit to Page' atau 'Scale to Fit'")
    print(f"   ✅ 2. AKTIFKAN 'Center Horizontally'")
    print(f"   ✅ 3. AKTIFKAN 'Center Vertically'")
    print(f"   ✅ 4. GUNAKAN margin minimal (5-10mm)")
    print(f"   ✅ 5. PASTIKAN orientasi kertas sesuai")
    
    print(f"\n⚙️  PENGATURAN OPTIMAL DI APLIKASI:")
    print(f"   📋 Paper Size: A4")
    print(f"   📐 Orientation: Portrait (atau sesuai kebutuhan)")
    print(f"   📏 Margins: 10mm (semua sisi)")
    print(f"   🔄 fit_to_page: true")
    print(f"   ↔️  center_horizontally: true")
    print(f"   ↕️  center_vertically: true")
    
    print(f"\n🎯 HASIL YANG DIHARAPKAN:")
    print(f"   📄 Dokumen akan diskalakan secara proporsional")
    print(f"   📍 Konten akan berada di tengah kertas")
    print(f"   📏 Margin merata di semua sisi")
    print(f"   ✅ Seluruh konten terlihat dan dapat dibaca")
    print(f"   🖨️  Hasil cetak memenuhi area kertas")
    
    print(f"\n🔍 CARA MENGECEK DI KODE APLIKASI:")
    print(f"   1. 📂 Buka file konfigurasi print settings")
    print(f"   2. ✅ Pastikan 'fit_to_page': true")
    print(f"   3. ✅ Pastikan 'center_horizontally': true")
    print(f"   4. ✅ Pastikan 'center_vertically': true")
    print(f"   5. 📏 Set margin ke 10mm atau kurang")
    
    print(f"\n🛠️  IMPLEMENTASI TEKNIS:")
    print(f"   1. 📐 Hitung rasio scaling berdasarkan ukuran kertas")
    print(f"   2. ⚖️  Gunakan scale factor terkecil (min(scale_x, scale_y))")
    print(f"   3. 📍 Hitung posisi tengah: (kertas - konten_scaled) / 2")
    print(f"   4. 🔄 Terapkan transformasi matrix untuk scaling dan positioning")
    print(f"   5. ✅ Validasi hasil sebelum mencetak")

def create_test_comparison():
    """
    Membuat perbandingan visual untuk demonstrasi
    """
    print(f"\n" + "="*60)
    print(f"📊 PERBANDINGAN VISUAL")
    print(f"="*60)
    
    print(f"\n❌ SEBELUM PERBAIKAN (Masalah):")
    print(f"   ┌─────────────────────┐")
    print(f"   │[DOC]                │  ← Dokumen di pojok")
    print(f"   │                     │")
    print(f"   │                     │")
    print(f"   │                     │")
    print(f"   │                     │")
    print(f"   │                     │")
    print(f"   │                     │")
    print(f"   │                     │")
    print(f"   └─────────────────────┘")
    print(f"   Sebagian besar area kosong, konten terpotong")
    
    print(f"\n✅ SETELAH PERBAIKAN (Solusi):")
    print(f"   ┌─────────────────────┐")
    print(f"   │                     │")
    print(f"   │   ┌─────────────┐   │")
    print(f"   │   │             │   │  ← Dokumen di tengah")
    print(f"   │   │   DOKUMEN   │   │     dengan scaling")
    print(f"   │   │             │   │")
    print(f"   │   └─────────────┘   │")
    print(f"   │                     │")
    print(f"   └─────────────────────┘")
    print(f"   Margin merata, seluruh konten terlihat")

if __name__ == "__main__":
    analyze_print_problem()
    provide_solutions()
    create_test_comparison()
    
    print(f"\n" + "="*60)
    print(f"🎉 ANALISIS SELESAI")
    print(f"="*60)
    print(f"\n📋 RINGKASAN:")
    print(f"   🔍 Masalah: Dokumen landscape dicetak di kertas portrait tanpa scaling")
    print(f"   💡 Solusi: Aktifkan fit_to_page + centering di aplikasi")
    print(f"   🎯 Hasil: Dokumen akan tercetak penuh dan terpusat di kertas")
    print(f"\n🚀 Langkah selanjutnya: Terapkan pengaturan optimal di aplikasi print")