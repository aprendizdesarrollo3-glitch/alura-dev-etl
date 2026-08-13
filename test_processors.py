import os

from processors import vacunacion, ambiente_superficie, aguas


SAMPLES_DIR = "sample_files"
OUTPUT_DIR = os.path.join("local_storage", "processed", "test")

os.makedirs(OUTPUT_DIR, exist_ok=True)


def test(nombre, func, filename):

    path = os.path.join(SAMPLES_DIR, filename)

    if not os.path.exists(path):
        print(f"[{nombre}] Archivo de prueba no encontrado: {path} (colocalo en sample_files/ para probar)")
        return

    try:
        resultados = func(path, OUTPUT_DIR)
        print(f"[{nombre}] OK -> {resultados}")

    except Exception as e:
        print(f"[{nombre}] ERROR -> {e}")


if __name__ == "__main__":
    test("vacunacion", vacunacion.procesar_archivo, "Vacunación Julio 2026.xlsx")
    test("ambiente_superficie", ambiente_superficie.procesar_archivo, "Ambiente y superficie Julio 2026.xlsx")
    test("aguas", aguas.procesar_archivo, "Aguas Julio 2026.xlsx")

    # Prueba de error intencional: archivo que no existe
    test("vacunacion (archivo inexistente)", vacunacion.procesar_archivo, "NoExiste.xlsx")