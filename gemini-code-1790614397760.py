def get_avatar_url(name, person_id):
    name_str = str(name).lower()
    
    # URL base apuntando a las imágenes subidas en tu repositorio actual
    raw_base = "https://raw.githubusercontent.com/bioast32-ai/Llamado-a-lista/principal"

    # Ashlin Matutes -> Ashlin.jpeg
    if "ashlin" in name_str or "matutes" in name_str:
        return f"{raw_base}/Ashlin.jpeg"

    # Deivyd Fonseca -> Deyvid.jpeg
    if "deivyd" in name_str or "deyvid" in name_str or "fonseca" in name_str:
        return f"{raw_base}/Deyvid.jpeg"

    # Nelson López -> Nelson.jpeg
    if "nelson" in name_str or "lopez" in name_str or "lópez" in name_str:
        return f"{raw_base}/Nelson.jpeg"

    # Yolanda Hernández (o Isabella) -> Isabella.jpeg
    if "yolanda" in name_str or "isabella" in name_str or "hernandez" in name_str or "hernández" in name_str:
        return f"{raw_base}/Isabella.jpeg"

    # Avatar genérico por IA para los demás
    return f"https://api.dicebear.com/7.x/avataaars/svg?seed={person_id}"