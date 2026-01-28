"""
Production Cleanup Script for Formatly
Removes development artifacts, cache files, logs, and build files
"""

import os
import shutil
from pathlib import Path

# Get the project root directory
PROJECT_ROOT = Path(__file__).parent

def delete_file(file_path):
    """Delete a single file"""
    try:
        if file_path.exists():
            file_path.unlink()
            print(f"✓ Deleted: {file_path.relative_to(PROJECT_ROOT)}")
            return True
        else:
            print(f"⊘ Not found: {file_path.relative_to(PROJECT_ROOT)}")
            return False
    except Exception as e:
        print(f"✗ Error deleting {file_path.relative_to(PROJECT_ROOT)}: {e}")
        return False

def delete_directory(dir_path):
    """Delete a directory and all its contents"""
    try:
        if dir_path.exists() and dir_path.is_dir():
            shutil.rmtree(dir_path)
            print(f"✓ Deleted directory: {dir_path.relative_to(PROJECT_ROOT)}")
            return True
        else:
            print(f"⊘ Directory not found: {dir_path.relative_to(PROJECT_ROOT)}")
            return False
    except Exception as e:
        print(f"✗ Error deleting {dir_path.relative_to(PROJECT_ROOT)}: {e}")
        return False

def main():
    print("=" * 60)
    print("🧹 FORMATLY PRODUCTION CLEANUP")
    print("=" * 60)
    
    deleted_count = 0
    
    # 1. Delete __pycache__ directories
    print("\n📁 Removing Python cache directories...")
    pycache_dirs = list(PROJECT_ROOT.rglob("__pycache__"))
    for pycache_dir in pycache_dirs:
        if delete_directory(pycache_dir):
            deleted_count += 1
    
    # 2. Delete .pyc files
    print("\n🐍 Removing .pyc files...")
    pyc_files = list(PROJECT_ROOT.rglob("*.pyc"))
    for pyc_file in pyc_files:
        if delete_file(pyc_file):
            deleted_count += 1
    
    # 3. Delete log files
    print("\n📝 Removing log files...")
    log_files = [
        PROJECT_ROOT / "app.log",
        PROJECT_ROOT / "results.log",
        PROJECT_ROOT / "debug_output.txt",
        PROJECT_ROOT / "desktop" / "theme_debug.log",
    ]
    for log_file in log_files:
        if delete_file(log_file):
            deleted_count += 1
    
    # Delete logs directory
    logs_dir = PROJECT_ROOT / "logs"
    if delete_directory(logs_dir):
        deleted_count += 1
    
    # 4. Delete test input/output files in misc
    print("\n🧪 Removing test files from misc...")
    test_files = [
        PROJECT_ROOT / "misc" / "test_input.docx",
        PROJECT_ROOT / "misc" / "test_output.docx",
    ]
    for test_file in test_files:
        if delete_file(test_file):
            deleted_count += 1
    
    # 5. Delete mock_api.py
    print("\n🔧 Removing mock API file...")
    mock_api = PROJECT_ROOT / "mock_api.py"
    if delete_file(mock_api):
        deleted_count += 1
    
    # 6. Delete build artifacts
    print("\n🏗️ Removing build artifacts...")
    build_artifacts = [
        PROJECT_ROOT / "Formatly.spec",
        PROJECT_ROOT / "dist",
        PROJECT_ROOT / "build",
        PROJECT_ROOT / "desktop" / "dist",
        PROJECT_ROOT / "desktop" / "build",
    ]
    for artifact in build_artifacts:
        if artifact.is_dir():
            if delete_directory(artifact):
                deleted_count += 1
        elif artifact.is_file():
            if delete_file(artifact):
                deleted_count += 1
    
    # Summary
    print("\n" + "=" * 60)
    print(f"✅ Cleanup complete! Removed {deleted_count} items.")
    print("=" * 60)
    print("\n📋 Items preserved:")
    print("  • formatting_stats.csv (kept as requested)")
    print("  • app_gemini.py (kept as requested)")
    print("  • app_huggingface.py (kept as requested)")
    print("  • misc/ folder (added to .gitignore)")
    print("  • tests/ folder (added to .gitignore)")
    print("  • documents/ folder (added to .gitignore)")
    print("\n⚠️  Note: .env files are now in .gitignore but not deleted.")
    print("   Make sure they're not committed to version control!")

if __name__ == "__main__":
    main()
