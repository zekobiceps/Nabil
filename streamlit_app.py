import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import io
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# ═══════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Générateur — Fin Période d'Essai",
    page_icon="📋",
    layout="wide",
)

# ═══════════════════════════════════════════════════════════
# FONCTIONS UTILITAIRES
# ═══════════════════════════════════════════════════════════

def parse_date(v):
    """Parse une valeur date de façon sécurisée."""
    if v is None:
        return None
    try:
        if isinstance(v, float) and pd.isna(v):
            return None
    except TypeError:
        pass
    try:
        return pd.to_datetime(v)
    except (ValueError, TypeError):
        return None


def get_gender_info(civ, default="Mme"):
    """Retourne (titre, is_female) selon la civilité."""
    if not civ or (isinstance(civ, float) and pd.isna(civ)):
        civ = default
    s = str(civ).upper().strip()
    if s in ("M.", "MR", "M", "MONSIEUR") or s.startswith("MR ") or s.startswith("M. "):
        return "M.", False
    if "MLLE" in s or "MADEMOISELLE" in s:
        return "Mlle", True
    return "Mme", True  # Mme par défaut


def calc_end_dates(date_entree, d1_days=45, d_total_months=3):
    """Calcule (date_fin_1ere_periode, date_titularisation)."""
    if date_entree is None:
        return None, None
    date_fin_1ere = date_entree + timedelta(days=d1_days)
    date_tit = date_entree + relativedelta(months=d_total_months) - timedelta(days=1)
    return date_fin_1ere, date_tit


def build_titularisation(nom, prenom, poste, date_entree, date_fin,
                          titre="Mme", is_female=True):
    """Génère le message de titularisation."""
    verb = "recrutée" if is_female else "recruté"
    return (
        "Bonjour ,\n\n"
        f"Je me permets de vous contacter au sujet de la titularisation de "
        f"{titre} {nom} {prenom}, {verb} en qualité de {poste} "
        f"depuis le {date_entree.strftime('%d/%m/%Y')}. "
        f"Sa période d'essai arrive à son terme le {date_fin.strftime('%d/%m/%Y')}.\n\n"
        "De ce fait, je vous prie de bien vouloir remplir le document ci-joint "
        "afin de confirmer ou non sa titularisation.\n\n"
        "Sincères salutations,"
    )


def build_prolongement(individu, nom, prenom, poste, direction, sup,
                        date_entree, date_fin_1ere, titre="Mme", is_female=True):
    """Génère le message de prolongement de période d'essai."""
    collab = "collaboratrice" if is_female else "collaborateur"
    hdr = "\t".join([
        titre, "NOM", "PRENOM", "FONCTION",
        "CHANTIER CTRL PRES", "SUP ", "DATE D'EMBAUCHE",
        "DATE FIN  1ERE PERIODE D'ESSAI"
    ])
    data_row = "\t".join([
        str(individu), nom, prenom, poste, direction, sup,
        date_entree.strftime("%Y-%m-%d"),
        date_fin_1ere.strftime("%d/%m/%Y"),
    ])
    return (
        "Bonjour , \n\n"
        f"Nous vous informons que la {collab} mentionné(e) ci-dessous "
        "arrive à la fin de leur première période d'essai.\n\n"
        f"{hdr}\n{data_row}\n\n"
        "Pouvez-vous nous confirmer si vous les considérez aptes à bénéficier "
        "d'une prolongation de leur période d'essai ?\n"
        "Nous vous prions de bien vouloir nous répondre par retour de mail.\n\n"
        "Cordialement, "
    )


def auto_map_columns(columns):
    """Détection automatique de la correspondance des colonnes."""
    mapping = {}
    for col in columns:
        cu = col.upper().replace(" ", "").replace("(", "").replace(")", "").strip()
        if cu == "NOM":
            mapping.setdefault("NOM", col)
        elif cu == "PRENOM":
            mapping.setdefault("PRENOM", col)
        elif cu in ("LIB", "LIBPOSTE", "POSTE", "FONCTION", "LIBELLEPOSTE"):
            mapping.setdefault("LIB", col)
        elif cu in ("LIB80", "LIB80DIRECTION", "DIRECTION", "CHANTIER", "CHANTIERCTRLPRES"):
            mapping.setdefault("LIB80", col)
        elif cu in ("SUP", "NOMDURESPONSABLE", "NOMDURESPHIERARCHIQUE", "NOMRESPONSABLE"):
            mapping.setdefault("SUP", col)
        elif "DATE" in cu and ("ENTREE" in cu or "EMBAUCHE" in cu):
            mapping.setdefault("DATE_ENTREE", col)
        elif cu in ("INDIVIDU", "MATRICULE", "MAT", "MLE", "MLLE"):
            mapping.setdefault("INDIVIDU", col)
        elif cu in ("CIVILITE", "TITRE", "CIVIL", "CIVILITÉ"):
            mapping.setdefault("CIVILITE", col)
        elif "EMAIL" in cu or "MAIL" in cu:
            mapping.setdefault("EMAIL", col)
    return mapping


def send_email_smtp(server, port, username, password, from_addr,
                     to_addr, subject, body, att_bytes=None, att_name=None):
    """Envoie un email via SMTP avec STARTTLS."""
    msg = MIMEMultipart()
    msg["From"] = from_addr
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))
    if att_bytes and att_name:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(att_bytes)
        encoders.encode_base64(part)
        safe_name = att_name.replace('"', "")
        part.add_header("Content-Disposition", f'attachment; filename="{safe_name}"')
        msg.attach(part)
    context = ssl.create_default_context()
    with smtplib.SMTP(server, int(port), timeout=30) as smtp:
        smtp.starttls(context=context)
        smtp.login(username, password)
        smtp.sendmail(from_addr, to_addr, msg.as_string())


def get_safe_str(row, col):
    """Retourne une chaîne propre depuis une cellule du DataFrame."""
    if col is None:
        return ""
    v = row.get(col, "")
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


# ═══════════════════════════════════════════════════════════
# SIDEBAR — CONFIGURATION
# ═══════════════════════════════════════════════════════════
with st.sidebar:
    st.title("⚙️ Configuration")

    msg_type = st.radio(
        "Type de message",
        ["📄 Titularisation", "🔄 Prolongement Période d'Essai"],
        index=0,
    )
    is_titularisation = msg_type.startswith("📄")

    st.divider()
    st.subheader("Durée période d'essai")
    d1_days = st.number_input(
        "1ère période (jours)", value=45, min_value=1, max_value=365,
        help="Nombre de jours pour la 1ère période d'essai (défaut : 45)"
    )
    d_total_months = st.number_input(
        "Durée totale (mois)", value=3, min_value=1, max_value=24,
        help="Durée totale avant titularisation (défaut : 3 mois)"
    )

    st.divider()
    default_civ = st.selectbox(
        "Civilité par défaut",
        ["Mme", "M.", "Mlle"],
        help="Utilisée si la colonne CIVILITE est absente du fichier"
    )

    st.divider()
    with st.expander("📧 Configuration Email SMTP"):
        smtp_server = st.text_input("Serveur SMTP", placeholder="smtp.gmail.com")
        smtp_port = st.number_input("Port", value=587, min_value=1, max_value=65535)
        smtp_user = st.text_input("Identifiant", placeholder="votre@email.com")
        smtp_pass = st.text_input("Mot de passe", type="password")
        smtp_from = st.text_input("Expéditeur (From)", placeholder="votre@email.com")

    email_configured = all([smtp_server, smtp_user, smtp_pass, smtp_from])

# ═══════════════════════════════════════════════════════════
# PAGE PRINCIPALE
# ═══════════════════════════════════════════════════════════
st.title("📋 Générateur — Fin de Période d'Essai")
mode_label = "Titularisation" if is_titularisation else "Prolongement Période d'Essai"
st.caption(f"Mode actif : **{mode_label}**")
st.divider()

# ── ÉTAPE 1 : Import ────────────────────────────────────────
st.subheader("① Importer la liste des collaborateurs")
st.caption("Colonnes attendues : NOM · PRENOM · LIB/POSTE · LIB80/DIRECTION · SUP · DATE ENTREE (+ optionnel : INDIVIDU · CIVILITE · EMAIL)")

uploaded = st.file_uploader(
    "Charger le fichier Excel ou CSV",
    type=["xlsx", "xls", "csv"],
    help="Format supporté : .xlsx, .xls, .csv"
)

if uploaded:
    try:
        if uploaded.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded)
        else:
            df = pd.read_excel(uploaded, engine="openpyxl" if uploaded.name.endswith("xlsx") else None)
        df = df.dropna(how="all").reset_index(drop=True)
    except Exception as e:
        st.error(f"❌ Impossible de lire le fichier : {e}")
        st.stop()

    st.success(f"✅ {len(df)} collaborateur(s) chargé(s)")

    with st.expander("📊 Aperçu du fichier importé", expanded=False):
        st.dataframe(df, use_container_width=True)

    # ── ÉTAPE 2 : Correspondance colonnes ───────────────────
    st.subheader("② Correspondance des colonnes")
    detected = auto_map_columns(df.columns.tolist())
    available = ["— Non défini —"] + df.columns.tolist()

    def col_select(label, key, required=True):
        default_val = detected.get(key, "— Non défini —")
        idx = available.index(default_val) if default_val in available else 0
        sel = st.selectbox(
            label + (" \\*" if required else " _(optionnel)_"),
            available,
            index=idx,
            key=f"colmap_{key}",
        )
        return None if sel == "— Non défini —" else sel

    c1, c2, c3 = st.columns(3)
    with c1:
        col_nom      = col_select("NOM", "NOM", required=True)
        col_prenom   = col_select("PRÉNOM", "PRENOM", required=True)
        col_lib      = col_select("POSTE / FONCTION", "LIB", required=True)
    with c2:
        col_lib80    = col_select("DIRECTION / CHANTIER", "LIB80", required=False)
        col_sup      = col_select("SUP (Responsable hiérarchique)", "SUP", required=False)
        col_date     = col_select("DATE D'ENTRÉE EN FONCTION", "DATE_ENTREE", required=True)
    with c3:
        col_individu = col_select("MATRICULE / INDIVIDU", "INDIVIDU", required=False)
        col_civ      = col_select("CIVILITÉ (M./Mme/Mlle)", "CIVILITE", required=False)
        col_email    = col_select("EMAIL destinataire", "EMAIL", required=False)

    required_ok = all([col_nom, col_prenom, col_lib, col_date])
    if not required_ok:
        st.error("⚠️ Les colonnes **NOM**, **PRÉNOM**, **POSTE** et **DATE D'ENTRÉE** sont obligatoires.")

    # ── ÉTAPE 3 : Pièce jointe ──────────────────────────────
    st.subheader("③ Pièce jointe email (optionnel)")

    ATTACHMENT_TITUL = "/workspaces/Nabil/FR EPE - HICHMINE Mohamed Topographe.xlsx"
    ATTACHMENT_PROLONG = "/workspaces/Nabil/Model PERIODE ESSAI NV.xlsx"

    if is_titularisation:
        default_att_path = ATTACHMENT_TITUL
        att_label = "Formulaire EPE (FR EPE…xlsx)"
    else:
        default_att_path = ATTACHMENT_PROLONG
        att_label = "Modèle Période d'Essai (Model PERIODE ESSAI NV.xlsx)"

    use_default_att = st.checkbox(
        f"Utiliser **{os.path.basename(default_att_path)}** comme pièce jointe",
        value=os.path.exists(default_att_path),
    )
    custom_att = None
    if not use_default_att:
        custom_att = st.file_uploader(
            "Charger une autre pièce jointe",
            type=["xlsx", "xls", "pdf", "doc", "docx"],
            key="custom_att",
        )

    def get_attachment():
        if custom_att:
            return custom_att.read(), custom_att.name
        if use_default_att and os.path.exists(default_att_path):
            with open(default_att_path, "rb") as f:
                return f.read(), os.path.basename(default_att_path)
        return None, None

    # ── ÉTAPE 4 : Génération ────────────────────────────────
    st.subheader("④ Générer les messages")

    if required_ok and st.button("🚀 Générer les messages", type="primary", use_container_width=True):
        messages, subjects, email_dests, errors = [], [], [], []

        for idx, row in df.iterrows():
            try:
                nom       = get_safe_str(row, col_nom).upper()
                prenom    = get_safe_str(row, col_prenom)
                poste     = get_safe_str(row, col_lib)
                direction = get_safe_str(row, col_lib80)
                sup       = get_safe_str(row, col_sup)
                individu  = get_safe_str(row, col_individu)
                email_dest = get_safe_str(row, col_email)

                civ_val   = row[col_civ] if col_civ and pd.notna(row.get(col_civ, None)) else default_civ
                titre, is_f = get_gender_info(civ_val, default_civ)

                date_entree = parse_date(row[col_date])
                if date_entree is None:
                    errors.append(f"Ligne {idx + 2} — date invalide pour {nom} {prenom}")
                    messages.append(""); subjects.append(""); email_dests.append("")
                    continue

                date_fin_1ere, date_tit = calc_end_dates(date_entree, d1_days, d_total_months)

                if is_titularisation:
                    msg  = build_titularisation(nom, prenom, poste, date_entree, date_tit, titre, is_f)
                    subj = f"Titularisation — {nom} {prenom} — {poste}"
                else:
                    msg  = build_prolongement(individu, nom, prenom, poste, direction, sup,
                                              date_entree, date_fin_1ere, titre, is_f)
                    subj = f"Prolongement Période d'Essai — {nom} {prenom}"

                messages.append(msg)
                subjects.append(subj)
                email_dests.append(email_dest)

            except Exception as e:
                errors.append(f"Ligne {idx + 2} — {e}")
                messages.append(""); subjects.append(""); email_dests.append("")

        st.session_state["messages"]   = messages
        st.session_state["subjects"]   = subjects
        st.session_state["email_dests"] = email_dests
        st.session_state["df_gen"]     = df
        st.session_state["gen_cols"]   = {
            "nom": col_nom, "prenom": col_prenom
        }
        if errors:
            st.warning("⚠️ Erreurs détectées :\n" + "\n".join(errors))

    # ── ÉTAPE 5 : Résultats ─────────────────────────────────
    if "messages" in st.session_state and st.session_state["messages"]:
        messages    = st.session_state["messages"]
        subjects    = st.session_state["subjects"]
        email_dests = st.session_state["email_dests"]
        df_gen      = st.session_state["df_gen"]
        gen_cols    = st.session_state["gen_cols"]

        valid = [(i, m, s, e)
                 for i, (m, s, e) in enumerate(zip(messages, subjects, email_dests))
                 if m]

        if not valid:
            st.error("Aucun message n'a pu être généré. Vérifiez le fichier importé.")
        else:
            st.success(f"✅ **{len(valid)} message(s) généré(s)**")

            # Bouton téléchargement global
            all_text = ""
            for i, msg, subj, _ in valid:
                row = df_gen.iloc[i]
                n = get_safe_str(row, gen_cols["nom"]).upper()
                p = get_safe_str(row, gen_cols["prenom"])
                all_text += f"{'=' * 60}\n{n} {p}\nObjet : {subj}\n{'=' * 60}\n{msg}\n\n"

            filename_dl = (
                f"messages_titularisation_{datetime.now().strftime('%Y%m%d')}.txt"
                if is_titularisation
                else f"messages_prolongement_{datetime.now().strftime('%Y%m%d')}.txt"
            )
            st.download_button(
                "⬇️ Télécharger tous les messages (.txt)",
                data=all_text.encode("utf-8"),
                file_name=filename_dl,
                mime="text/plain",
                use_container_width=True,
            )

            st.divider()
            st.subheader("⑤ Messages générés")

            for i, msg, subj, email_dest in valid:
                row = df_gen.iloc[i]
                n = get_safe_str(row, gen_cols["nom"]).upper()
                p = get_safe_str(row, gen_cols["prenom"])
                label = f"👤  {n} {p}"
                if email_dest:
                    label += f"  —  {email_dest}"

                with st.expander(label, expanded=False):
                    edited = st.text_area(
                        "✏️ Message (modifiable avant envoi)",
                        value=msg,
                        height=300,
                        key=f"msg_edit_{i}",
                    )
                    st.caption(f"**Objet proposé :** {subj}")

                    if email_configured and email_dest:
                        if st.button(f"📤 Envoyer cet email", key=f"send_one_{i}"):
                            try:
                                att_b, att_n = get_attachment()
                                send_email_smtp(
                                    smtp_server, smtp_port, smtp_user, smtp_pass,
                                    smtp_from, email_dest, subj, edited, att_b, att_n,
                                )
                                st.success("✅ Email envoyé avec succès")
                            except Exception as e:
                                st.error(f"❌ Erreur envoi : {e}")
                    elif email_dest and not email_configured:
                        st.info("ℹ️ Configurez le SMTP dans la barre latérale pour envoyer par email.")
                    elif not email_dest:
                        st.caption("_Aucune adresse email renseignée pour ce collaborateur._")

            # Envoi groupé
            if email_configured:
                st.divider()
                recipients_with_email = [(i, m, s, e) for i, m, s, e in valid if e]
                if recipients_with_email:
                    st.subheader("📤 Envoi groupé")
                    st.caption(
                        f"{len(recipients_with_email)} destinataire(s) avec adresse email trouvé(s)"
                    )
                    if st.button(
                        f"📤 Envoyer tous les emails ({len(recipients_with_email)})",
                        type="primary",
                        use_container_width=True,
                    ):
                        att_b, att_n = get_attachment()
                        sent_ok, sent_fail = 0, []
                        prog = st.progress(0)
                        for step, (i, msg, subj, email_dest) in enumerate(recipients_with_email):
                            try:
                                send_email_smtp(
                                    smtp_server, smtp_port, smtp_user, smtp_pass,
                                    smtp_from, email_dest, subj, msg, att_b, att_n,
                                )
                                sent_ok += 1
                            except Exception as e:
                                sent_fail.append(f"{email_dest} : {e}")
                            prog.progress((step + 1) / len(recipients_with_email))

                        if sent_ok:
                            st.success(f"✅ {sent_ok} email(s) envoyé(s)")
                        if sent_fail:
                            st.error("Erreurs :\n" + "\n".join(sent_fail))

else:
    # ── Page d'accueil (aucun fichier) ──────────────────────
    st.info(
        """
        **Comment utiliser cet outil :**

        1. Choisissez le **type de message** dans la barre latérale
        2. **Importez** votre fichier Excel / CSV avec les colonnes collaborateurs
        3. Vérifiez la **correspondance des colonnes**
        4. Cliquez sur **Générer les messages**
        5. **Téléchargez** le fichier texte ou **envoyez** directement par email

        ---
        **Colonnes supportées dans votre fichier :**

        | Colonne | Description | Obligatoire |
        |---------|-------------|-------------|
        | NOM | Nom de famille | ✅ |
        | PRENOM | Prénom | ✅ |
        | LIB / POSTE | Intitulé du poste | ✅ |
        | DATE ENTREE | Date d'entrée en fonction | ✅ |
        | LIB80 / DIRECTION | Direction / Chantier | — |
        | SUP | Nom du responsable hiérarchique | — |
        | INDIVIDU / MAT | Matricule | — |
        | CIVILITE | M. / Mme / Mlle | — |
        | EMAIL | Email du destinataire | — |
        """
    )
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 📄 Exemple — Titularisation")
        st.code(
            "Bonjour ,\n\n"
            "Je me permets de vous contacter au sujet de la\n"
            "titularisation de Mme NOM PRENOM, recrutée en\n"
            "qualité de POSTE depuis le JJ/MM/AAAA.\n"
            "Sa période d'essai arrive à son terme le JJ/MM/AAAA.\n\n"
            "De ce fait, je vous prie de bien vouloir remplir le\n"
            "document ci-joint afin de confirmer ou non sa titularisation.\n\n"
            "Sincères salutations,",
            language=None,
        )
    with col2:
        st.markdown("#### 🔄 Exemple — Prolongement")
        st.code(
            "Bonjour ,\n\n"
            "Nous vous informons que la collaboratrice mentionnée\n"
            "ci-dessous arrive à la fin de leur première période d'essai.\n\n"
            "Mme  NOM  PRENOM  FONCTION  CHANTIER  SUP  DATE  FIN 1ERE\n"
            "MAT  ...  ...     ...       ...        ...  ...   ...\n\n"
            "Pouvez-vous nous confirmer si vous les considérez aptes\n"
            "à bénéficier d'une prolongation de leur période d'essai ?\n"
            "Nous vous prions de bien vouloir nous répondre par retour de mail.\n\n"
            "Cordialement,",
            language=None,
        )

