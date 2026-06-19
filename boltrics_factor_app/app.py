import io
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Boltrics Celbezetting", layout="wide")
BASE = Path(__file__).parent
MAP_FILE = BASE / "pallet_mapping.csv"
CEL_FILE = BASE / "cel_instellingen.csv"

st.title("Boltrics palletlocaties overzicht")
st.caption("Upload de Handling Unit export. Pas palletfactoren en celinstellingen zelf aan in de zijbalk.")

@st.cache_data
def load_mapping():
    df = pd.read_csv(MAP_FILE, dtype={"Code": str})
    df["Code"] = df["Code"].str.strip().str.upper()
    df["Factor"] = pd.to_numeric(df["Factor"], errors="coerce").fillna(0.0)
    return df

@st.cache_data
def load_cells():
    df = pd.read_csv(CEL_FILE, dtype={"Cel": str})
    df["Cel"] = df["Cel"].str.zfill(2)
    df["Capaciteit"] = pd.to_numeric(df["Capaciteit"], errors="coerce").fillna(0)
    return df

mapping_default = load_mapping()
cells_default = load_cells()

with st.sidebar:
    st.header("Instellingen")
    st.subheader("Palletfactoren")
    st.write("Pas de Factor aan. Nieuwe codes kun je toevoegen onderaan de tabel.")
    mapping_edit = st.data_editor(
        mapping_default,
        num_rows="dynamic",
        use_container_width=True,
        column_config={"Factor": st.column_config.NumberColumn("Factor", min_value=0.0, step=0.05, format="%.2f")},
        key="mapping_editor",
    )
    st.download_button(
        "Download pallet_mapping.csv",
        mapping_edit.to_csv(index=False).encode("utf-8"),
        "pallet_mapping.csv",
        "text/csv",
    )
    st.divider()
    st.subheader("Celinstellingen")
    cells_edit = st.data_editor(
        cells_default,
        num_rows="dynamic",
        use_container_width=True,
        column_config={"Capaciteit": st.column_config.NumberColumn("Capaciteit", min_value=0, step=1)},
        key="cells_editor",
    )
    st.download_button(
        "Download cel_instellingen.csv",
        cells_edit.to_csv(index=False).encode("utf-8"),
        "cel_instellingen.csv",
        "text/csv",
    )

uploaded = st.file_uploader("Upload Boltrics Handling Unit Excel", type=["xlsx", "xls"])

if not uploaded:
    st.info("Upload een Handling Unit export om het dashboard te tonen.")
    st.stop()

try:
    raw = pd.read_excel(uploaded, dtype=str)
except Exception as e:
    st.error(f"Excel kon niet gelezen worden: {e}")
    st.stop()

if raw.shape[1] < 3:
    st.error("De export heeft minimaal 3 kolommen nodig: A=HU, B=Location No., C=Handling Unit Type Code.")
    st.stop()

hu_col = raw.columns[0]
loc_col = raw.columns[1]
type_col = raw.columns[2]

df = raw.copy()
df["HU"] = df[hu_col].astype(str).str.strip()
df["Location No."] = df[loc_col].astype(str).str.strip()
df["Handling Unit Type Code"] = df[type_col].astype(str).str.strip().str.upper()
df = df[(df["Location No."].notna()) & (df["Location No."] != "") & (df["Location No."].str.lower() != "nan")]
df["Cel"] = df["Location No."].str[:2].str.zfill(2)

map_df = mapping_edit.copy()
map_df["Code"] = map_df["Code"].astype(str).str.strip().str.upper()
map_df["Factor"] = pd.to_numeric(map_df["Factor"], errors="coerce").fillna(0.0)
factor_dict = dict(zip(map_df["Code"], map_df["Factor"]))
df["Factor"] = df["Handling Unit Type Code"].map(factor_dict).fillna(0.0)

summary = df.groupby("Cel").agg(
    **{"Gewogen bezet": ("Factor", "sum"), "HU's": ("HU", "count"), "Unieke locaties": ("Location No.", "nunique")}
).reset_index()
summary["Gestapeld"] = summary["HU's"] - summary["Unieke locaties"]
summary["Stapeling %"] = (summary["Gestapeld"] / summary["HU's"]).fillna(0)

cells = cells_edit.copy()
cells["Cel"] = cells["Cel"].astype(str).str.zfill(2)
cells["Capaciteit"] = pd.to_numeric(cells["Capaciteit"], errors="coerce").fillna(0)
summary = cells.merge(summary, on="Cel", how="left")
for col in ["Gewogen bezet", "HU's", "Unieke locaties", "Gestapeld", "Stapeling %"]:
    summary[col] = summary[col].fillna(0)
summary["Vrij"] = summary["Capaciteit"] - summary["Gewogen bezet"]
summary["Bezetting %"] = (summary["Gewogen bezet"] / summary["Capaciteit"]).where(summary["Capaciteit"] > 0, 0)

cols = ["Cel", "Inhoud", "Vrij", "Gewogen bezet", "Temperatuur", "Capaciteit", "Bezetting %", "HU's", "Unieke locaties", "Gestapeld", "Stapeling %"]
summary = summary[cols]

st.subheader("Celbezetting overzicht")
st.dataframe(
    summary.style.format({
        "Vrij": "{:.2f}",
        "Gewogen bezet": "{:.2f}",
        "Bezetting %": "{:.1%}",
        "Stapeling %": "{:.1%}",
    }),
    use_container_width=True,
    hide_index=True,
)

st.subheader("Cel detail")
selected = st.selectbox("Kies cel", summary["Cel"].tolist())
detail = df[df["Cel"] == selected].copy()
detail["Dubbele locatie"] = detail.duplicated("Location No.", keep=False)
st.dataframe(
    detail[["HU", "Location No.", "Handling Unit Type Code", "Factor", "Dubbele locatie"]].sort_values(["Location No.", "HU"]),
    use_container_width=True,
    hide_index=True,
)

double_locs = detail[detail["Dubbele locatie"]].sort_values(["Location No.", "HU"])
with st.expander("Dubbel gebruikte locaties in deze cel"):
    if double_locs.empty:
        st.success("Geen dubbele locaties gevonden in deze cel.")
    else:
        st.dataframe(double_locs[["HU", "Location No.", "Handling Unit Type Code", "Factor"]], use_container_width=True, hide_index=True)

output = io.BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    summary.to_excel(writer, sheet_name="Cel_Overzicht", index=False)
    df.to_excel(writer, sheet_name="HU_Detail", index=False)
    map_df.to_excel(writer, sheet_name="Pallet_Mapping", index=False)
    cells.to_excel(writer, sheet_name="Cel_Instellingen", index=False)

st.download_button("Download resultaat naar Excel", output.getvalue(), "boltrics_celbezetting_resultaat.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
