import base64
import os
import requests

_VALID_EXTS = {".xlsx", ".xls", ".pdf"}


def _share_id(share_url: str) -> str:
    encoded = base64.urlsafe_b64encode(share_url.encode()).rstrip(b"=").decode()
    return f"u!{encoded}"


def _headers(token: str | None) -> dict:
    if token:
        return {"Authorization": f"Bearer {token}"}
    return {}


def _list_children(share_url: str, token: str | None = None) -> list[dict]:
    sid = _share_id(share_url)
    headers = _headers(token)
    endpoints = [
        f"https://graph.microsoft.com/v1.0/shares/{sid}/driveItem/children"
        "?$select=name,file,@microsoft.graph.downloadUrl",
        f"https://api.onedrive.com/v1.0/shares/{sid}/root/children",
    ]
    for url in endpoints:
        try:
            r = requests.get(url, headers=headers, timeout=20)
            if r.ok:
                return r.json().get("value", [])
        except Exception:
            continue
    raise RuntimeError(
        "No se pudo listar la carpeta. Verifica que el enlace sea público "
        "o autentica tu cuenta Microsoft primero."
    )


def download_fichas_from_onedrive(
    share_url: str, dest_dir: str, token: str | None = None
) -> tuple[int, list[str]]:
    """
    Descarga archivos de fichas desde un share link de OneDrive.
    Retorna (n_descargados, [nombres_descargados]).
    """
    os.makedirs(dest_dir, exist_ok=True)
    items = _list_children(share_url, token=token)
    headers = _headers(token)
    downloaded, names = 0, []

    for item in items:
        name = item.get("name", "")
        if os.path.splitext(name)[1].lower() not in _VALID_EXTS:
            continue
        dl_url = item.get("@microsoft.graph.downloadUrl")
        if not dl_url:
            continue
        content = requests.get(dl_url, headers=headers, timeout=60).content
        with open(os.path.join(dest_dir, name), "wb") as f:
            f.write(content)
        downloaded += 1
        names.append(name)

    return downloaded, names
