import sys
import io
import traceback
from colorama import init as _colorama_init, Fore
from asana_outlook_integration_script import main as run

_colorama_init()

class BinaryWriter:
    """Captures text and streams it live as green 8-bit binary, one bit per line."""
    def __init__(self, real_stream):
        self.real = real_stream
        self.buffer = io.StringIO()

    def write(self, text: str):
        # record for later replay
        self.buffer.write(text)
        for ch in text:
            bits = format(ord(ch), "08b")
            for bit in bits:
                self.real.write(Fore.GREEN + bit + "\n")
        self.real.flush()

    def flush(self):
        self.real.flush()


def main_wrapper():
    real_out = sys.stdout
    bw = BinaryWriter(real_out)
    sys.stdout = sys.stderr = bw

    try:
        print("Starting main.py execution\n")
        run()
        print("\nScript completed\n")
    except Exception:
        traceback.print_exc()
    finally:
        sys.stdout = sys.stderr = real_out

    print("\n=== Original output below ===\n")
    print(bw.buffer.getvalue())

if __name__ == "__main__":
    main_wrapper()
