
#!/usr/bin/env python3
"""
Server Integration for Enhanced Print Service
Integrasi enhanced print service dengan server yang ada
"""

import sys
import os
from pathlib import Path

# Add current directory to path
sys.path.insert(0, str(Path(__file__).parent))

from enhanced_print_service import EnhancedPrintService

def patch_main_server():
    """Patch main server untuk menggunakan enhanced print service"""
    
    main_file = Path("server/main.py")
    
    if not main_file.exists():
        print(f"❌ Main server file not found: {main_file}")
        return False
    
    # Backup original file
    backup_file = main_file.with_suffix(".py.backup")
    if not backup_file.exists():
        import shutil
        shutil.copy2(main_file, backup_file)
        print(f"✓ Backup created: {backup_file}")
    
    # Read current content
    with open(main_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Add enhanced print service import
    if "from enhanced_print_service import EnhancedPrintService" not in content:
        import_line = "from enhanced_print_service import EnhancedPrintService
"
        
        # Find a good place to insert the import
        lines = content.split('
')
        insert_index = 0
        
        for i, line in enumerate(lines):
            if line.startswith('import ') or line.startswith('from '):
                insert_index = i + 1
        
        lines.insert(insert_index, import_line.strip())
        content = '
'.join(lines)
    
    # Replace print service usage
    replacements = [
        # Replace DirectPrintService with EnhancedPrintService
        ("DirectPrintService()", "EnhancedPrintService()"),
        ("print_service = DirectPrintService()", "print_service = EnhancedPrintService()"),
        
        # Replace print method calls
        (".print_file_direct(", ".print_pdf_with_fallbacks("),
        (".print_pdf(", ".print_pdf_with_fallbacks("),
    ]
    
    for old, new in replacements:
        if old in content:
            content = content.replace(old, new)
            print(f"✓ Replaced: {old} -> {new}")
    
    # Write updated content
    with open(main_file, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"✓ Server integration completed: {main_file}")
    return True

def test_server_integration():
    """Test server integration"""
    print("
=== TESTING SERVER INTEGRATION ===")
    
    try:
        # Test import
        from enhanced_print_service import EnhancedPrintService
        print("✓ Enhanced print service import successful")
        
        # Test service creation
        service = EnhancedPrintService()
        print("✓ Enhanced print service creation successful")
        
        # Test printer finding
        printer = service.find_printer()
        if printer:
            print(f"✓ Printer found: {printer}")
        else:
            print("⚠️  No printer found")
        
        return True
        
    except Exception as e:
        print(f"❌ Integration test failed: {e}")
        return False

if __name__ == "__main__":
    if patch_main_server():
        test_server_integration()
