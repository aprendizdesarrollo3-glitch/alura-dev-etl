import os


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)
    return path


def save_bytes(data, dest_path):
    ensure_dir(os.path.dirname(dest_path))

    with open(dest_path, "wb") as f:
        f.write(data)

    return dest_path