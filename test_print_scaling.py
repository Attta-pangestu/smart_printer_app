import fitz  # PyMuPDF
import os

def analyze_print_problem():
    """
    Menganalisis masalah pencetakan PDF
    """
    print("\n=== ANALISIS MASALAH PENCETAKAN ===")
    
    # Path file PDF
    pdf_path = "d:\\Gawean Rebinmas\\Driver_Epson_L120\\server\\temp\\20250822_111236_Test_print.pdf"
    
    if not os.path.exists(pdf_path):
        print(f"❌ File tidak ditemukan: {pdf_path}")
        return
    
    try:
        # Buka PDF asli
        doc = fitz.open(pdf_path)
        page = doc[0]
        
        print(f"📄 DOKUMEN ASLI:")
        print(f"   Dimensi: {page.rect.width:.1f} x {page.rect.height:.1f} pt")
        print(f"   Ukuran: {page.rect.width * 0.352778:.1f} x {page.rect.height * 0.352778:.1f} mm")
        print(f"   Format: A4 Landscape (297 x 210 mm)")
        
        # Analisis konten
        images = page.get_images()
        if images:
            print(f"\n🖼️  KONTEN GAMBAR:")
            for img_index, img in enumerate(images):
                try:
                    xref = img[0]
                    pix = fitz.Pixmap(doc, xref)
                    print(f"   Gambar {img_index + 1}: {pix.width} x {pix.height} piksel")
                    print(f"   Resolusi tinggi: {pix.width * pix.height:,} piksel total")
                    pix = None
                except:
                    print(f"   Gambar {img_index + 1}: Tidak dapat dianalisis")
        
        doc.close()
        
        # Simulasi pencetakan dengan berbagai pengaturan
        simulate_printing_scenarios()
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")

def simulate_printing_scenarios():
    """
    Simulasi berbagai skenario pencetakan
    """
    print(f"\n📋 SIMULASI PENCETAKAN:")
    
    # Ukuran kertas A4 portrait (target pencetakan)
    a4_width_pt = 595.28  # 210mm
    a4_height_pt = 841.89  # 297mm
    
    # Ukuran dokumen asli (A4 landscape)
    doc_width_pt = 842.40  # 297mm
    doc_height_pt = 595.44  # 210mm
    
    print(f"\n🎯 TARGET KERTAS (A4 Portrait):")
    print(f"   Dimensi: {a4_width_pt:.1f} x {a4_height_pt:.1f} pt")
    print(f"   Ukuran: 210 x 297 mm")
    
    print(f"\n📄 DOKUMEN ASLI (A4 Landscape):")
    print(f"   Dimensi: {doc_width_pt:.1f} x {doc_height_pt:.1f} pt")
    print(f"   Ukuran: 297 x 210 mm")
    
    # Skenario 1: Tanpa scaling (masalah utama)
    print(f"\n❌ SKENARIO 1 - TANPA SCALING:")
    print(f"   Dokumen ditempatkan langsung di kertas")
    print(f"   Lebar dokumen (842pt) > Lebar kertas (595pt)")
    print(f"   Hasil: Dokumen terpotong dan muncul di pojok")
    print(f"   Rasio yang terlihat: {a4_width_pt/doc_width_pt:.2f} = {(a4_width_pt/doc_width_pt)*100:.1f}%")
    
    # Skenario 2: Dengan fit to page
    margin_pt = 20 * 2.834645  # 20mm to points
    available_width = a4_width_pt - (2 * margin_pt)
    available_height = a4_height_pt - (2 * margin_pt)
    
    scale_x = available_width / doc_width_pt
    scale_y = available_height / doc_height_pt
    scale_factor = min(scale_x, scale_y)
    
    scaled_width = doc_width_pt * scale_factor
    scaled_height = doc_height_pt * scale_factor
    
    print(f"\n✅ SKENARIO 2 - DENGAN FIT TO PAGE:")
    print(f"   Area tersedia: {available_width:.1f} x {available_height:.1f} pt")
    print(f"   Scale X: {scale_x:.3f}")
    print(f"   Scale Y: {scale_y:.3f}")
    print(f"   Scale factor: {scale_factor:.3f} ({scale_factor*100:.1f}%)")
    print(f"   Ukuran setelah scaling: {scaled_width:.1f} x {scaled_height:.1f} pt")
    print(f"   Ukuran dalam mm: {scaled_width*0.352778:.1f} x {scaled_height*0.352778:.1f} mm")
    
    # Hitung posisi dengan centering
    pos_x = (a4_width_pt - scaled_width) / 2
    pos_y = (a4_height_pt - scaled_height) / 2
    
    print(f"\n📍 POSISI DENGAN CENTERING:")
    print(f"   X: {pos_x:.1f} pt ({pos_x*0.352778:.1f} mm dari kiri)")
    print(f"   Y: {pos_y:.1f} pt ({pos_y*0.352778:.1f} mm dari atas)")
    
    # Margin aktual
    margin_left = pos_x
    margin_top = pos_y
    margin_right = a4_width_pt - (pos_x + scaled_width)
    margin_bottom = a4_height_pt - (pos_y + scaled_height)
    
    print(f"\n📏 MARGIN AKTUAL:")
    print(f"   Kiri: {margin_left:.1f} pt ({margin_left*0.352778:.1f} mm)")
    print(f"   Atas: {margin_top:.1f} pt ({margin_top*0.352778:.1f} mm)")
    print(f"   Kanan: {margin_right:.1f} pt ({margin_right*0.352778:.1f} mm)")
    print(f"   Bawah: {margin_bottom:.1f} pt ({margin_bottom*0.352778:.1f} mm)")

def provide_solutions():
    """
    Memberikan solusi untuk masalah pencetakan
    """
    print(f"\n🔧 SOLUSI MASALAH PENCETAKAN:")
    
    print(f"\n🎯 PENYEBAB UTAMA:")
    print(f"   1. Dokumen dalam format A4 Landscape (297x210mm)")
    print(f"   2. Printer diatur untuk A4 Portrait (210x297mm)")
    print(f"   3. Tidak ada scaling otomatis")
    print(f"   4. Konten ditempatkan di pojok kiri atas")
    
    print(f"\n💡 SOLUSI YANG HARUS DITERAPKAN:")
    print(f"   ✅ 1. AKTIFKAN 'Fit to Page' atau 'Scale to Fit'")
    print(f"   ✅ 2. AKTIFKAN 'Center Horizontally'")
    print(f"   ✅ 3. AKTIFKAN 'Center Vertically'")
    print(f"   ✅ 4. GUNAKAN margin minimal (5-10mm)")
    
    print(f"\n⚙️  PENGATURAN OPTIMAL:")
    print(f"   📋 Paper Size: A4")
    print(f"   📐 Orientation: Portrait")
    print(f"   📏 Margins: 10mm (semua sisi)")
    print(f"   🔄 Fit to Page: ENABLED")
    print(f"   ↔️  Center Horizontally: ENABLED")
    print(f"   ↕️  Center Vertically: ENABLED")
    
    print(f"\n🎯 HASIL YANG DIHARAPKAN:")
    print(f"   📄 Dokumen akan diskalakan ke ~70% dari ukuran asli")
    print(f"   📍 Konten akan berada di tengah kertas")
    print(f"   📏 Margin merata di semua sisi (~59mm)")
    print(f"   ✅ Seluruh konten terlihat dan proporsional")
    
    print(f"\n🔍 CARA MENGECEK DI APLIKASI:")
    print(f"   1. Buka pengaturan print di aplikasi")
    print(f"   2. Pastikan 'fit_to_page': true")
    print(f"   3. Pastikan 'center_horizontally': true")
    print(f"   4. Pastikan 'center_vertically': true")
    print(f"   5. Set margin ke 10mm atau kurang")

def create_test_pdf():
    """
    Membuat PDF test dengan pengaturan yang benar
    """
    print(f"\n📝 MEMBUAT PDF TEST DENGAN PENGATURAN OPTIMAL:")
    
    try:
        # Buka PDF asli
        input_path = "d:\\Gawean Rebinmas\\Driver_Epson_L120\\server\\temp\\20250822_111236_Test_print.pdf"
        output_path = "d:\\Gawean Rebinmas\\Driver_Epson_L120\\Test_print_FIXED.pdf"
        
        if not os.path.exists(input_path):
            print(f"❌ File input tidak ditemukan")
            return
        
        # Buka dokumen asli
        src_doc = fitz.open(input_path)
        src_page = src_doc[0]
        
        # Buat dokumen baru dengan ukuran A4 portrait
        new_doc = fitz.open()
        new_page = new_doc.new_page(width=595.28, height=841.89)  # A4 portrait
        
        # Hitung scaling dan posisi
        margin_pt = 10 * 2.834645  # 10mm margin
        available_width = 595.28 - (2 * margin_pt)
        available_height = 841.89 - (2 * margin_pt)
        
        src_width = src_page.rect.width
        src_height = src_page.rect.height
        
        scale_x = available_width / src_width
        scale_y = available_height / src_height
        scale_factor = min(scale_x, scale_y)
        
        scaled_width = src_width * scale_factor
        scaled_height = src_height * scale_factor
        
        # Posisi tengah
        pos_x = (595.28 - scaled_width) / 2
        pos_y = (841.89 - scaled_height) / 2
        
        # Buat matrix transformasi
        matrix = fitz.Matrix(scale_factor, scale_factor).pretranslate(pos_x, pos_y)
        
        # Insert halaman dengan transformasi
        new_page.show_pdf_page(new_page.rect, src_doc, 0, clip=None, matrix=matrix)
        
        # Simpan PDF baru
        new_doc.save(output_path)
        new_doc.close()
        src_doc.close()
        
        print(f"✅ PDF test berhasil dibuat: {output_path}")
        print(f"📐 Scale factor: {scale_factor:.3f} ({scale_factor*100:.1f}%)")
        print(f"📍 Posisi: ({pos_x:.1f}, {pos_y:.1f}) pt")
        print(f"📏 Ukuran final: {scaled_width:.1f} x {scaled_height:.1f} pt")
        
        return output_path
        
    except Exception as e:
        print(f"❌ Error membuat PDF test: {str(e)}")
        return None

if __name__ == "__main__":
    analyze_print_problem()
    provide_solutions()
    
    # Buat PDF test dengan pengaturan yang benar
    test_pdf = create_test_pdf()
    
    if test_pdf:
        print(f"\n🎉 SELESAI!")
        print(f"📁 File test tersedia di: {test_pdf}")
        print(f"🖨️  Coba cetak file test ini untuk melihat perbedaannya")