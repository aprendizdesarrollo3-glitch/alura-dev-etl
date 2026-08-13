from processors import vacunacion, ambiente_superficie, aguas


# Mapa nombre_logico -> funcion procesar_archivo(input_file, output_dir)
PROCESSORS = {
    "vacunacion": vacunacion.procesar_archivo,
    "ambiente_superficie": ambiente_superficie.procesar_archivo,
    "aguas": aguas.procesar_archivo,
}


def get_processor(filename):
    """
    Decide que procesador usar segun el nombre del archivo adjunto.

    Ejemplos reales:
      "Vacunacion Julio 2026.xlsx"           -> vacunacion
      "Vacunación Julio 2026.xlsx"           -> vacunacion
      "Ambiente y superficie Julio 2026.xlsx" -> ambiente_superficie
      "Aguas Julio 2026.xlsx"                -> aguas
    """
    name = filename.lower()

    if "vacun" in name:
        return PROCESSORS["vacunacion"]

    if "ambiente" in name or "superficie" in name:
        return PROCESSORS["ambiente_superficie"]

    if "agua" in name:
        return PROCESSORS["aguas"]

    return None