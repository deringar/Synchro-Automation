import time
from pynput import keyboard
from pynput.keyboard import Controller, Key

# Create keyboard controller
kb_controller = Controller()

# Set key to press and interval (in seconds)
key_to_press = Key.shift
interval = 10  # seconds

# Flag to control the loop
running = True

# Function to handle key press events
def on_press(key):
    global running
    if key == Key.esc:
        print("Escape key pressed. Exiting...")
        running = False
        return False  # Stop the listener

# Start the key listener in a non-blocking way
listener = keyboard.Listener(on_press=on_press)
listener.start()

print("Running. Press Esc to stop.")
try:
    while running:
        kb_controller.press(key_to_press)
        kb_controller.release(key_to_press)
        print(f"Pressed '{key_to_press}'")
        time.sleep(interval)
except Exception as e:
    print("Error:", e)

listener.join()
