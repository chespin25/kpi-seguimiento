import os
import streamlit as st
from supabase import create_client, Client

@st.cache_resource
def get_client() -> Client:
    url = os.environ.get("SUPABASE_URL") or st.secrets.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_SERVICE_KEY") or st.secrets.get("SUPABASE_SERVICE_KEY")
    if not url or not key:
        st.error("Faltan credenciales SUPABASE_URL y SUPABASE_SERVICE_KEY")
        st.stop()
    return create_client(url, key)
