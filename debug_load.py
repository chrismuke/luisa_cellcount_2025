import numpy as np
import sys
import pickle

def investigate_file(filepath):
    """
    Attempts to load and inspect a .npy file to understand its structure.
    """
    print(f"--- Investigating: {filepath} ---")

    # Attempt 1: Standard numpy load with allow_pickle=True
    try:
        data = np.load(filepath, allow_pickle=True)
        print("Successfully loaded with np.load.")
        print(f"  - Type: {type(data)}")
        print(f"  - Shape: {getattr(data, 'shape', 'N/A')}")
        print(f"  - Dtype: {getattr(data, 'dtype', 'N/A')}")

        if getattr(data, 'ndim', -1) == 0:
            print("  - It's a 0-dim array, attempting to extract item...")
            item = data.item()
            print(f"  - Extracted item type: {type(item)}")
            if isinstance(item, dict):
                print(f"  - Item is a dict. Keys: {item.keys()}")
            else:
                print(f"  - Item is not a dict. Content: {item}")
        return
    except Exception as e:
        print(f"np.load failed: {e}")

    # Attempt 2: If np.load fails, try to open as a raw binary file
    print("\nAttempting to read as a raw file to check for pickle format...")
    try:
        with open(filepath, 'rb') as f:
            # The first few bytes can identify a numpy file or a pickle file
            magic_bytes = f.read(8)
            print(f"  - First 8 bytes: {magic_bytes}")
            if magic_bytes.startswith(b'\x93NUMPY'):
                print("  - This looks like a standard numpy array file.")
            elif magic_bytes.startswith(b'\x80'):
                print("  - This might be a pickle file. Trying to load with pickle.")
                f.seek(0) # Rewind to start
                p_data = pickle.load(f)
                print("  - Successfully loaded with pickle.load().")
                print(f"  - Data type: {type(p_data)}")
                if isinstance(p_data, dict):
                    print(f"  - Keys: {p_data.keys()}")

    except Exception as e:
        print(f"Raw file read or pickle.load failed: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        investigate_file(sys.argv[1])
    else:
        print("Please provide a file path to investigate.")