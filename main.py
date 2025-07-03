import sys, io, traceback
from colorama import init as _colorama_init, Fore
from asana_outlook_integration_script import main as run

_colorama_init()

class BinaryWriter:
    """Wraps a real stream, writing every incoming chunk as green 8-bit binary."""
    def __init__(self, real):
        self.real = real
    def write(self, s):
        if not s:
            return
        # Convert each character to binary and output in green
        bits = " ".join(format(ord(ch), "08b") for ch in s)
        self.real.write(Fore.GREEN + bits + "\n")
    def flush(self):
        self.real.flush()

def main_wrapper():
    # Swap out stdout/stderr for live binary streaming
    old_out, old_err = sys.stdout, sys.stderr
    bin_out = BinaryWriter(old_out)
    sys.stdout = sys.stderr = bin_out
    try:
        print("Starting main.py execution")
        run()
        print("Script completed")
    except Exception:
        traceback.print_exc()
    finally:
        sys.stdout, sys.stderr = old_out, old_err

    # restore real streams, capture traceback if any
    sys.stdout, sys.stderr = old_out, old_err
    print("\n\n=== Run complete. Final logs / traceback below ===\n")
    # Now replay the captured text so you can read it
    print(bin_out.real.getvalue() if hasattr(bin_out.real, "getvalue") else "")

if __name__ == "__main__":
    main_wrapper()
