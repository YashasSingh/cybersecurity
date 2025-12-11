"""
Auto-generated automation script from click recording
Generated: 2025-12-11 02:34:42
Recorded 6 clicks
"""
import time
from pywinauto import Application
from pywinauto.mouse import click as mouse_click

def main():
    print("🤖 Starting RustDesk automation...")
    
    # Connect to RustDesk
    app = Application(backend="uia").connect(title_re=".*RustDesk.*", timeout=10)
    window = app.window(title_re=".*RustDesk.*")
    window.set_focus()
    time.sleep(0.5)
    
    print(f"✓ Connected to: {window.window_text()}")
    
    # Click 1: Pane
    print("Click 1: Pane")
    try:
        control = window.descendants()[0]
        control.click_input()
    except Exception as e:
        print(f"  ⚠ Failed to click control [0]: {e}")
        # Fallback: Click by coordinates
        mouse_click(coords=(960, 504))
    time.sleep(0.5)  # Wait between clicks
    
    # Click 2: Static - "Security"
    print("Click 2: Static")
    try:
        control = window.descendants()[16]
        control.click_input()
    except Exception as e:
        print(f"  ⚠ Failed to click control [16]: {e}")
        # Fallback: Click by coordinates
        mouse_click(coords=(147, 241))
    time.sleep(0.5)  # Wait between clicks
    
    # Click 3: Button - "Set permanent password"
    print("Click 3: Button")
    try:
        control = window.descendants()[62]
        control.click_input()
    except Exception as e:
        print(f"  ⚠ Failed to click control [62]: {e}")
        # Fallback: Click by coordinates
        mouse_click(coords=(584, 745))
    time.sleep(0.5)  # Wait between clicks
    
    # Click 4: GroupBox - "Enable file transfer"
    print("Click 4: GroupBox")
    try:
        control = window.descendants()[34]
        control.click_input()
    except Exception as e:
        print(f"  ⚠ Failed to click control [34]: {e}")
        # Fallback: Click by coordinates
        mouse_click(coords=(703, 363))
    time.sleep(0.5)  # Wait between clicks
    
    # Click 5: GroupBox - "Enable TCP tunneling"
    print("Click 5: GroupBox")
    try:
        control = window.descendants()[42]
        control.click_input()
    except Exception as e:
        print(f"  ⚠ Failed to click control [42]: {e}")
        # Fallback: Click by coordinates
        mouse_click(coords=(703, 555))
    time.sleep(0.5)  # Wait between clicks
    
    # Click 6: Button
    print("Click 6: Button")
    try:
        control = window.descendants()[91]
        control.click_input()
    except Exception as e:
        print(f"  ⚠ Failed to click control [91]: {e}")
        # Fallback: Click by coordinates
        mouse_click(coords=(0, 0))
    time.sleep(0.5)  # Wait between clicks
    
    print("✅ Automation complete!")

if __name__ == "__main__":
    main()
