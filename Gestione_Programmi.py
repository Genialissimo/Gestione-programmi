"""
Gestione Programmi - Web App (Streamlit + Google Sheets)

Scheletro costruito sulla stessa architettura di "Gestione Registrazioni SEG":
sidebar con logout, navigazione tramite card cliccabili (session_state.pagina),
stesso stile CSS per card/tab/post-it.
"""

from datetime import datetime

import pandas as pd
import streamlit as st

import gspread
from google.oauth2.service_account import Credentials

# ==============================================================================
# 1. CONFIGURAZIONE PAGINA (Deve essere la prima istruzione Streamlit)
# ==============================================================================
st.set_page_config(
    page_title="Gestione Programmi",
    page_icon="🗓️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────
# COSTANTI E CONFIGURAZIONI DEL SISTEMA
# ─────────────────────────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

NOME_FOGLIO_UTENTI = "Utenti"
RIGA_INTESTAZIONE_UTENTI = 1

NOME_FOGLIO_IMPEGNI = "Impegni"
NOME_FOGLIO_SETTIMANA = "Programmazione Settimanale"
NOME_FOGLIO_ADUNANZE = "Adunanze"
NOME_FOGLIO_MINISTERO = "Ministero"
NOME_FOGLIO_COMUNICAZIONI = "Comunicazioni"
NOME_FOGLIO_ANNUNCI = "Annunci"


# ─────────────────────────────────────────────────────────────────
# CONNESSIONE A GOOGLE SHEETS
# ─────────────────────────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def get_client() -> gspread.Client:
    """Autentica il programma verso Google tramite l'account di servizio."""
    credenziali = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=SCOPES
    )
    return gspread.authorize(credenziali)


@st.cache_resource(show_spinner=False)
def apri_foglio_dati():
    """Apre il foglio Google dati."""
    try:
        client = get_client()
        wb = client.open_by_key(st.secrets["sheet_id"])
        return wb, None
    except gspread.exceptions.APIError:
        email_sa = st.secrets["gcp_service_account"]["client_email"]
        return None, (
            "Impossibile aprire il foglio dati. Controlla che sia stato "
            f"condiviso (come Editor) con:\n`{email_sa}`"
        )
    except Exception as e:
        return None, f"Errore durante il collegamento: {e}"


@st.cache_data(ttl=60, show_spinner=False)
def leggi_foglio_come_df(_workbook, nome_foglio: str, riga_intestazione: int = 1):
    """Legge un foglio (tab) del workbook e lo ritorna come DataFrame."""
    try:
        ws = _workbook.worksheet(nome_foglio)
    except gspread.WorksheetNotFound:
        nomi_disponibili = ", ".join(f"'{f.title}'" for f in _workbook.worksheets())
        return None, (
            f"Il foglio '{nome_foglio}' non esiste nel documento collegato. "
            f"Fogli disponibili: {nomi_disponibili}."
        )
    except Exception as e:
        return None, f"Errore durante la lettura del foglio: {e}"

    tutti_i_valori = ws.get_all_values()
    if len(tutti_i_valori) < riga_intestazione:
        return pd.DataFrame(), None

    intestazioni = tutti_i_valori[riga_intestazione - 1]
    righe_dati = tutti_i_valori[riga_intestazione:]

    intestazioni_pulite = []
    contatori = {}
    for i, nome in enumerate(intestazioni):
        nome = nome.strip() or f"Colonna {i + 1}"
        if nome in contatori:
            contatori[nome] += 1
            nome = f"{nome} ({contatori[nome]})"
        else:
            contatori[nome] = 0
        intestazioni_pulite.append(nome)

    righe_dati = [r for r in righe_dati if any(cella.strip() for cella in r)]

    df = pd.DataFrame(righe_dati, columns=intestazioni_pulite)
    return df, None


def salva_riga_foglio(_workbook, nome_foglio: str, riga_intestazione: int,
                       valori: dict, riga_da_aggiornare: int = None):
    """Scrive o aggiorna una riga nel foglio."""
    try:
        ws = _workbook.worksheet(nome_foglio)
        intestazioni = ws.row_values(riga_intestazione)
        riga_completa = [valori.get(nome, "") for nome in intestazioni]

        if riga_da_aggiornare is None:
            ws.append_row(riga_completa, value_input_option="USER_ENTERED")
        else:
            ultima_colonna = gspread.utils.rowcol_to_a1(1, len(intestazioni)).rstrip("0123456789")
            intervallo = f"A{riga_da_aggiornare}:{ultima_colonna}{riga_da_aggiornare}"
            ws.update(intervallo, [riga_completa], value_input_option="USER_ENTERED")
        return True, None
    except Exception as e:
        return False, f"Errore durante il salvataggio: {e}"


def elimina_riga_foglio(_workbook, nome_foglio: str, riga_da_eliminare: int):
    """Elimina una riga da un foglio."""
    try:
        ws = _workbook.worksheet(nome_foglio)
        ws.delete_rows(riga_da_eliminare)
        return True, None
    except Exception as e:
        return False, f"Errore durante l'eliminazione: {e}"


def leggi_utente_da_email(_workbook, email: str):
    """Cerca l'email verificata da Google nel foglio 'Utenti' controllando
    le colonne specifiche: B (Utente), C (Indirizzo), D (Ruolo)."""
    df, err = leggi_foglio_come_df(_workbook, NOME_FOGLIO_UTENTI, RIGA_INTESTAZIONE_UTENTI)
    if err or df is None or df.empty:
        return None, None

    colonne_lower = {str(c).strip().lower(): c for c in df.columns}
    
    col_email = colonne_lower.get("indirizzo") or colonne_lower.get("email")
    col_nome = colonne_lower.get("utente") or colonne_lower.get("cognome e nome")
    col_ruolo = colonne_lower.get("ruolo")

    # Fallback posizionale se le intestazioni differiscono: B=1, C=2, D=3
    if not col_email and len(df.columns) >= 3:
        col_nome = df.columns[1]   # Colonna B
        col_email = df.columns[2]  # Colonna C
        if len(df.columns) >= 4:
            col_ruolo = df.columns[3] # Colonna D

    if not col_email:
        return None, None

    email_norm = (email or "").strip().lower()
    
    # Scansione dinamica su tutto il DataFrame
    corrispondenza = df[df[col_email].astype(str).str.strip().str.lower() == email_norm]
    
    if corrispondenza.empty:
        return None, None

    riga = corrispondenza.iloc[0]
    nome = str(riga.get(col_nome, "")).strip() if col_nome else ""
    ruolo_grezzo = str(riga.get(col_ruolo, "")).strip().lower() if col_ruolo else "utente"
    
    if ruolo_grezzo not in ("amministratore", "editor", "utente"):
        ruolo_grezzo = "utente"

    return (nome or email), ruolo_grezzo


# ==============================================================================
# 2. PANNELLO DI AUTENTICAZIONE
# ==============================================================================
if not st.user.is_logged_in:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("🔒 Accesso Riservato")
        st.subheader("Gestione Programmi")
        st.write("Accedi con il tuo account Google per entrare nell'applicazione.")
        st.login()  # Chiamata diretta senza st.button
    st.stop()


# ==============================================================================
# 3. AREA RISERVATA
# ==============================================================================
workbook, errore = apri_foglio_dati()
collegato = workbook is not None

if "ruolo" not in st.session_state or st.session_state.get("email_verificata") != st.user.email:
    if not collegato:
        st.error("⚠️ Impossibile verificare l'utente: il foglio dati non è raggiungibile.")
        st.caption(errore or "")
        if st.button("🚪 Esci", use_container_width=True):
            st.logout()
        st.stop()

    nome_trovato, ruolo_trovato = leggi_utente_da_email(workbook, st.user.email)
    if not nome_trovato:
        st.error(f"⚠️ L'indirizzo **{st.user.email}** non è autorizzato ad accedere a questa "
                 f"applicazione. Contatta l'amministratore per farti aggiungere al foglio «{NOME_FOGLIO_UTENTI}».")
        if st.button("🚪 Esci", use_container_width=True):
            st.logout()
        st.stop()

    st.session_state.nome_utente = nome_trovato
    st.session_state.ruolo = ruolo_trovato
    st.session_state.email_verificata = st.user.email

with st.sidebar:
    st.write("👤 Utente connesso:")
    st.write(f"**{st.session_state.nome_utente}**")
    st.caption(f"Ruolo: {st.session_state.ruolo.capitalize()}")
    st.caption(f"📧 `{st.user.email}`")
    if st.button("🚪 Logout", type="secondary", use_container_width=True):
        for chiave in ("nome_utente", "ruolo", "email_verificata"):
            st.session_state.pop(chiave, None)
        st.logout()

# ─────────────────────────────────────────────────────────────────
# NAVIGAZIONE
# ─────────────────────────────────────────────────────────────────
if "pagina" not in st.session_state:
    st.session_state.pagina = "home"


def vai_a(pagina: str):
    st.session_state.pagina = pagina


CARD_INFORMAZIONI = [
    ("", "", "Adunanze", "Informazioni e materiale relativo alle adunanze.", "info_adunanze"),
    ("", "", "Ministero", "Informazioni e materiale relativo al ministero.", "info_ministero"),
    ("", "", "Comunicazioni", "Comunicazioni della congregazione.", "info_comunicazioni"),
    ("", "", "Annunci", "Annunci correnti.", "info_annunci"),
]


def mostra_griglia_card(lista_card):
    for i in range(0, len(lista_card), 2):
        coppia = lista_card[i:i + 2]
        cols = st.columns(2)
        for col, (icon, bg_cls, titolo, desc, pagina) in zip(cols, coppia):
            with col:
                with st.container(key=f"card_{pagina}", border=True):
                    if icon:
                        intestazione_html = f"""
                        <div class="custom-card-header">
                            <div class="custom-card-title-group">
                                <span class="custom-icon-box {bg_cls}">{icon}</span>
                                <span class="custom-card-title">{titolo}</span>
                            </div>
                        </div>
                        """
                    else:
                        intestazione_html = f"""
                        <div class="custom-card-header">
                            <div class="custom-card-title-group">
                                <span class="custom-card-title">{titolo}</span>
                            </div>
                        </div>
                        """
                    st.markdown(intestazione_html, unsafe_allow_html=True)
                    st.caption(desc)

                    st.button(" ", key=f"nav_{pagina}", disabled=not collegato,
                              on_click=vai_a, args=(pagina,), use_container_width=True)


def _inietta_css_home():
    """CSS Custom per card, tab e post-it."""
    st.markdown("""
    <style>
        .custom-card-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            margin-bottom: 8px;
            width: 100%;
        }
        .custom-card-title-group {
            display: flex;
            align-items: center;
            gap: 10px;
            flex: 1;
            min-width: 0;
        }
        .custom-card-title {
            font-size: 1.1rem !important;
            font-weight: 700 !important;
            line-height: 1.3 !important;
            margin: 0 !important;
            padding: 0 !important;
        }
        .custom-icon-box {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 38px;
            height: 38px;
            min-width: 38px;
            border-radius: 10px;
            font-size: 1.3rem;
            flex-shrink: 0;
        }
        .bg-orange { background: rgba(249, 115, 22, 0.15); border: 1px solid rgba(249, 115, 22, 0.4); }
        .bg-blue   { background: rgba(59, 130, 246, 0.15); border: 1px solid rgba(59, 130, 246, 0.4); }
        .bg-green  { background: rgba(16, 185, 129, 0.15); border: 1px solid rgba(16, 185, 129, 0.4); }
        .bg-purple { background: rgba(168, 85, 247, 0.15); border: 1px solid rgba(168, 85, 247, 0.4); }
        .bg-cyan   { background: rgba(6, 182, 212, 0.15); border: 1px solid rgba(6, 182, 212, 0.4); }
        .bg-amber  { background: rgba(245, 158, 11, 0.15); border: 1px solid rgba(245, 158, 11, 0.4); }
        .bg-slate  { background: rgba(100, 116, 139, 0.15); border: 1px solid rgba(100, 116, 139, 0.4); }

        .hud-badge {
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 0.8rem;
            font-weight: 600;
            white-space: nowrap;
            display: inline-block;
        }
        .hud-green {
            background: rgba(16, 185, 129, 0.15);
            color: #10b981;
            border: 1px solid rgba(16, 185, 129, 0.3);
        }
        .hud-yellow {
            background: rgba(245, 158, 11, 0.15);
            color: #f59e0b;
            border: 1px solid rgba(245, 158, 11, 0.3);
        }
        .hud-red {
            background: rgba(239, 68, 68, 0.15);
            color: #ef4444;
            border: 1px solid rgba(239, 68, 68, 0.3);
        }

        .stTabs [data-baseweb="tab-list"] {
            gap: 6px;
            border-bottom: 1px solid rgba(128,128,128,0.3);
        }
        .stTabs [data-baseweb="tab"] {
            font-weight: 600;
            font-size: 0.95rem;
            color: var(--text-color);
            opacity: 0.65;
            padding: 10px 6px;
        }
        .stTabs [aria-selected="true"] {
            color: #2E7D32 !important;
            opacity: 1 !important;
            font-weight: 700 !important;
        }
        .stTabs [data-baseweb="tab-highlight"] {
            background-color: #2E7D32 !important;
            height: 3px !important;
            border-radius: 2px;
        }

        div[class*="st-key-card_"] {
            position: relative !important;
            box-shadow: 3px 5px 14px rgba(0,0,0,0.18);
            transition: border-color 0.15s ease, box-shadow 0.15s ease, transform 0.15s ease;
            cursor: pointer;
        }
        div[class*="st-key-card_"]:hover {
            border-color: #2E7D32 !important;
            box-shadow: 4px 7px 18px rgba(0,0,0,0.24);
            transform: translateY(-2px);
        }

        div[class*="st-key-card_"] .custom-card-header,
        div[class*="st-key-card_"] [data-testid="stCaptionContainer"] {
            pointer-events: none !important;
        }

        div[class*="st-key-card_"] div[data-testid="stElementContainer"]:has(div[data-testid="stButton"]) {
            position: absolute !important;
            inset: 0 !important;
            width: 100% !important;
            height: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
            z-index: 5 !important;
        }

        div[class*="st-key-card_"] div[data-testid="stButton"] {
            width: 100% !important;
            height: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
        }

        div[class*="st-key-card_"] div[data-testid="stButton"] button {
            width: 100% !important;
            height: 100% !important;
            opacity: 0 !important;
            background: transparent !important;
            border: none !important;
            cursor: pointer !important;
            margin: 0 !important;
            padding: 0 !important;
        }

        div[class*="st-key-card_"] div[data-testid="stButton"] button:disabled {
            cursor: not-allowed !important;
        }

        .postit-card {
            width: 92%;
            max-width: 900px;
            margin: 8px auto 24px auto;
            background: linear-gradient(135deg, #fff9c4, #fff3a0);
            border-radius: 4px 14px 4px 14px;
            padding: 22px clamp(20px, 4vw, 40px);
            box-shadow: 3px 5px 14px rgba(0,0,0,0.18);
            transform: rotate(-0.8deg);
        }
        .postit-titolo {
            font-size: clamp(1.05rem, 1.6vw, 1.3rem);
            font-weight: 700;
            color: #5c4a00;
            margin: 0 0 8px 0;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .postit-testo {
            font-size: 0.95rem;
            color: #4a3f00;
            line-height: 1.4;
        }
    </style>
    """, unsafe_allow_html=True)


def mostra_home():
    ora_ora = datetime.now().strftime('%d/%m/%Y %H:%M')
    st.markdown(
        f"""
        <div style="margin-bottom: 12px;">
            <h3 style="font-size: 1.25rem; font-weight: 700; margin: 0; padding: 0;">🗓️ Gestione Programmi</h3>
            <p style="font-size: 0.8rem; color: #6b7280; margin: 2px 0 0 0; padding: 0;">
                Ultimo aggiornamento: {ora_ora}
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    _inietta_css_home()

    postit_html = """
    <div class="postit-card">
        <div class="postit-titolo">📌 Benvenuto/a</div>
        <div class="postit-testo">
            Da qui puoi gestire i tuoi impegni, la programmazione della settimana
            e consultare le informazioni di congregazione.
        </div>
    </div>
    """

    nomi_tab = ["🏠 Home", "📋 I miei impegni", "📅 Questa settimana", "📊 Tabella Informazioni", "🗓️ Programmi"]
    tabs = st.tabs(nomi_tab)

    with tabs[0]:
        st.markdown(postit_html, unsafe_allow_html=True)

    with tabs[1]:
        st.subheader("📋 I miei impegni")
        st.caption("🚧 Sezione in costruzione.")

    with tabs[2]:
        st.subheader("📅 Questa settimana")
        st.caption("🚧 Sezione in costruzione.")

    with tabs[3]:
        mostra_griglia_card(CARD_INFORMAZIONI)

    with tabs[4]:
        st.subheader("🗓️ Programmi")
        st.caption("🚧 Sezione in costruzione.")


def _pagina_segnaposto(titolo: str, emoji: str, nome_foglio_futuro: str):
    st.title(f"{emoji} {titolo}")
    st.button("🏠 Torna alla Home", key=f"home_da_{titolo.lower()}", use_container_width=True,
              on_click=vai_a, args=("home",))

    if not collegato:
        st.warning("⚠️ Nessun foglio dati collegato.")
        return

    st.info(f"🚧 Sezione in costruzione. Questa pagina leggerà i dati dal foglio Google «{nome_foglio_futuro}».")


def mostra_info_adunanze():
    _pagina_segnaposto("Adunanze", "🙌", NOME_FOGLIO_ADUNANZE)


def mostra_info_ministero():
    _pagina_segnaposto("Ministero", "📖", NOME_FOGLIO_MINISTERO)


def mostra_info_comunicazioni():
    _pagina_segnaposto("Comunicazioni", "📢", NOME_FOGLIO_COMUNICAZIONI)


def mostra_info_annunci():
    _pagina_segnaposto("Annunci", "📣", NOME_FOGLIO_ANNUNCI)


def mostra_tabella_informazioni_ridotta():
    ora_ora = datetime.now().strftime('%d/%m/%Y %H:%M')
    st.markdown(
        f"""
        <div style="margin-bottom: 12px;">
            <h3 style="font-size: 1.25rem; font-weight: 700; margin: 0; padding: 0;">📊 Tabella Informazioni</h3>
            <p style="font-size: 0.8rem; color: #6b7280; margin: 2px 0 0 0; padding: 0;">
                Ultimo aggiornamento: {ora_ora}
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )
    _inietta_css_home()
    mostra_griglia_card(CARD_INFORMAZIONI)


# ─────────────────────────────────────────────────────────────────
# CONTROLLO ACCESSO E ROUTING
# ─────────────────────────────────────────────────────────────────
PAGINE_CONSENTITE_UTENTE = {"home", "info_adunanze", "info_ministero", "info_comunicazioni", "info_annunci"}
if st.session_state.ruolo == "utente" and st.session_state.pagina not in PAGINE_CONSENTITE_UTENTE:
    st.session_state.pagina = "home"

if st.session_state.ruolo == "utente" and st.session_state.pagina == "home":
    mostra_tabella_informazioni_ridotta()
    st.stop()

if st.session_state.pagina == "info_adunanze":
    mostra_info_adunanze()
elif st.session_state.pagina == "info_ministero":
    mostra_info_ministero()
elif st.session_state.pagina == "info_comunicazioni":
    mostra_info_comunicazioni()
elif st.session_state.pagina == "info_annunci":
    mostra_info_annunci()
else:
    mostra_home()
