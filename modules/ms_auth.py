import time
import requests
import streamlit as st

_TOKEN_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
_DEVICE_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode"
_SCOPES = "https://graph.microsoft.com/Files.Read offline_access"


def start_device_flow(client_id: str, tenant_id: str) -> dict:
    resp = requests.post(
        _DEVICE_URL.format(tenant=tenant_id),
        data={"client_id": client_id, "scope": _SCOPES},
        timeout=15,
    )
    resp.raise_for_status()
    flow = resp.json()
    st.session_state["ms_flow"] = flow
    st.session_state["ms_client_id"] = client_id
    st.session_state["ms_tenant_id"] = tenant_id
    return flow


def poll_for_token() -> tuple[str | None, str]:
    """Intenta obtener el token una vez. Retorna (token, mensaje)."""
    flow      = st.session_state.get("ms_flow", {})
    client_id = st.session_state.get("ms_client_id", "")
    tenant_id = st.session_state.get("ms_tenant_id", "")
    if not flow or not client_id:
        return None, "Inicia el flujo primero."

    if time.time() > flow.get("expires_in", 0) + flow.get("_ts", time.time()):
        st.session_state.pop("ms_flow", None)
        return None, "El código expiró. Inicia el flujo de nuevo."

    resp = requests.post(
        _TOKEN_URL.format(tenant=tenant_id),
        data={
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": client_id,
            "device_code": flow["device_code"],
        },
        timeout=15,
    )
    data = resp.json()

    if "access_token" in data:
        expiry = time.time() + int(data.get("expires_in", 3600))
        st.session_state["ms_token"] = data["access_token"]
        st.session_state["ms_token_expiry"] = expiry
        st.session_state.pop("ms_flow", None)
        return data["access_token"], "Autenticado correctamente."

    error = data.get("error", "")
    if error == "authorization_pending":
        return None, "Esperando... autentícate en el navegador y vuelve a hacer clic."
    if error == "slow_down":
        return None, "Demasiados intentos, espera 10 segundos."
    return None, f"Error: {data.get('error_description', error)}"


def get_valid_token() -> str | None:
    token  = st.session_state.get("ms_token")
    expiry = st.session_state.get("ms_token_expiry", 0)
    if token and time.time() < expiry - 60:
        return token
    return None


def clear_token():
    for k in ("ms_token", "ms_token_expiry", "ms_flow", "ms_client_id", "ms_tenant_id"):
        st.session_state.pop(k, None)
