import sys
import io
import traceback
from colorama import init as _colorama_init, Fore
from asana_outlook_integration_script import main as run

_colorama_init()

def main_wrapper():
    # Capture all stdout/stderr
    old_out, old_err = sys.stdout, sys.stderr
    buf = io.StringIO()
    sys.stdout, sys.stderr = buf, buf
    try:
        print("Starting main.py execution")
        run()
        print("Script completed")
    except Exception:
        buf.write(traceback.format_exc())
    finally:
        sys.stdout, sys.stderr = old_out, old_err

    output = buf.getvalue()
    # Convert each character to 8-bit binary
    binary = " ".join(format(ord(c), "08b") for c in output)
    # Print binary in green, then original text
    print(Fore.GREEN + binary)
    print(output)

if __name__ == "__main__":
    main_wrapper()
