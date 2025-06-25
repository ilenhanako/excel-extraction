#!/usr/bin/env python3
"""
Launcher script to run the apps
"""

import sys
import os

def main():
    """Main launcher function."""
    print("🚀 Excel Data Extraction Gradio Apps")
    print("=" * 50)
    print("Choose which app to run:")
    print("1. Basic App (app.py) - Simple extraction and visualization")
    print("2. Enhanced App (app_enhanced.py) - Advanced features and detailed analysis")
    print("3. Exit")
    
    while True:
        choice = input("\nEnter your choice (1-3): ").strip()
        
        if choice == "1":
            print("\n📊 Launching Basic Excel Extraction App...")
            try:
                import app
                print("✅ Basic app launched successfully!")
            except ImportError as e:
                print(f"❌ Error importing basic app: {e}")
                print("Make sure all dependencies are installed: pip install -r requirements.txt")
            except Exception as e:
                print(f"❌ Error running basic app: {e}")
            break
            
        elif choice == "2":
            print("\n🚀 Launching Enhanced Excel Extraction App...")
            try:
                import app_enhanced
                print("✅ Enhanced app launched successfully!")
            except ImportError as e:
                print(f"❌ Error importing enhanced app: {e}")
                print("Make sure all dependencies are installed: pip install -r requirements.txt")
            except Exception as e:
                print(f"❌ Error running enhanced app: {e}")
            break
            
        elif choice == "3":
            print("👋 Goodbye!")
            sys.exit(0)
            
        else:
            print("❌ Invalid choice. Please enter 1, 2, or 3.")

if __name__ == "__main__":
    main() 