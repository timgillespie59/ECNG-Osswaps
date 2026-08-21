"""
ECNG Supply — Mobile Snapshot v2
Same styling/branding/data as the current app, restructured behind a
home screen with three sections: Outstanding, Transacted, Pricing.
"""
import json
from datetime import datetime, date
from pathlib import Path

import streamlit as st

import excel_parser
import onedrive_source

st.set_page_config(page_title="ECNG Supply", page_icon="⚡", layout="centered")

NAVY = "#002F6C"
GOLD = "#FFCD00"
GREEN = "#43B02A"
GRAY = "#989898"
DARK_GRAY = "#333333"
RED = "#D6483F"

STATUS_COLORS = {"Active": NAVY, "In the Money": GREEN, "Expired": GRAY, "Transacted": GOLD}

WATERMARK_B64 = "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"

LOGO_B64 = "iVBORw0KGgoAAAANSUhEUgAAABEAAAAcCAYAAACH81QkAAAEPklEQVR4nKWUXWxUVRDH/zPn7vZzXQrYGhpDGoJI12jLQpWmZRMJip8Jxrs+NKIm2MaaggXbojG5ezU+SKJt3Whs9yohkcRs9ckqCEFcjIqGDw2pJgiIRjCk8lFpu93uvWd82LYsAiHBebuTe35nzvznP4T/E3duKikskoeNGzstBJD4CyYq3Ul+gm+AQIjEFAC4rqojouEbgcglGuqznuzhS9+g61YAINiwuQzlP0s43OzTIreuqL758DRE8m/IPzQTEUsBoNF0utM4Xll/wh9YJBlv+LUnF1YZADD7nvZKX4AvnNn95vg1oSloAKI9bw0T5qZdOYYJb0/W1TnIRZfK5VzWAtCMiGWgPCQVF34vPLO7Y2yKSQDpwtrW+VnQMECLXC1347ex5X+dz7QwALz8UOAnz5Woql2/BSnbDZw8NOvs2T/ehWUxYPGUGjTJ6kFiHtQixyA4CXw4djHtzmXAVLZtaybsAPAM17Q9p40LSqvitb5PRx4HbI1hMACBpnuJ9C6DeV9BAfcAIEVUxlh9i4GIZQB0mpheIUZ9Olv6mZLMG57oDYFlz89BKOQV3dVeySRl2ZLZR24q9H+y9L7gPgCiFFUydsYzSNkuFH8pWi98bEHj0wTt95HRw8DHYy69j4Gol1XuWoC+Rsp2z35bPJaybTeZTPoBVLA//EKnEd7YBkiDQB4YGIh6VXOzdePKG3EPv90NUECF21/VoBU+v3IAkGmGCABGR0fni4jHWntprb1uaGgiGqxYtank2M54BvvfmgCEZvmwRgRNTHpHen/3KZgmm2ZOM8/zlgA4xe7heJygPyCSU96heNeZydJMbtBIAMK5H+L/3D7x52L3YDwOgDAw4A0NDREAEFEDEX1HCDf7gm5RyUWWA8TS6h2M70LEMpCy3bzJvWKak8kkj4yMfAWgeSbrW7r+Dq5p+7WovmNeLmPlm3PGAslkUokI9ff3N/b39+8FAC6t21RthDvrKgx1nBnvZNIT3/hr2qovTeqMDQAAQ0NDRERCRM8S0SAAMMgVLZktpydlHMTLQeq4KI5GImAgdpkJLctiANpxnHkAGrXW2wCAR7/v/aWW0qsI0g3BnBLFrdmDvbHUTE8kH8S2bWsReZ2IBltaWv62LMug3NttDQBqyYb7BbKZiH80CPHMgZ4T+b2IRqOe4zi1IvJ5cXFx9dGjR0disZhMAwgRy/AO9X6hHylbycRHPNBL/mXtXYH6jkWmaappWQFsF5Gupqam86FQiIjov8pdpggKwi8uCDRsvM2yrEIAcBznPcdxtk/1Z2bJX30lmqbCQFIDJEyAFiCRSLQAeCoYDDbmfjH1Vaq4Mpr7+nxTFTyaSCT29fX1BQGQiFxvH089LCcnEonEasdxPtq6deus/Px1Y/omx3FWOo7TNWX5awKuKEtEKBaLUVVV1WLXdcvXrVu3dzp/rR78C8ki27gtR+ssAAAAAElFTkSuQmCC"

# ---------------------------------------------------------------- styles
st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@600;700&display=swap');

    .stApp {{
        background: linear-gradient(180deg, #eef2f8 0%, #fbfcfe 22%, #ffffff 55%);
    }}
    .block-container {{ padding-top: 1rem; padding-bottom: 3rem; max-width: 480px; }}
    .ecng-header {{
        position: relative; background-color: {NAVY}; padding: 16px 20px 14px 20px;
        border-radius: 8px 8px 0 0; margin-bottom: 0; overflow: hidden;
    }}
    .ecng-header-row {{ display: flex; align-items: center; gap: 12px; position: relative; z-index: 1; }}
    .ecng-logo-badge {{
        width: 40px; height: 40px; border-radius: 50%; background: #fff;
        display: flex; align-items: center; justify-content: center; flex-shrink: 0;
        box-shadow: 0 2px 6px rgba(0,0,0,0.25);
    }}
    .ecng-logo-badge img {{ max-width: 24px; max-height: 28px; width: auto; height: auto; }}
    .ecng-header h1 {{ color: #fff; margin: 0; font-size: 1.35rem; font-family: 'Poppins', sans-serif; }}
    .ecng-header p {{ color: #b9c9e2; margin: 2px 0 0 0; font-size: 0.8rem; }}
    .ecng-header-divider {{ height: 3px; background-color: {GOLD}; }}
    .ecng-header-divider.rounded {{ border-radius: 0 0 8px 8px; margin-bottom: 18px; }}

    .ecng-ticker {{
        background-color: {NAVY}; border-radius: 0 0 8px 8px; margin-top: 0; margin-bottom: 6px;
        padding: 12px 20px 10px 20px; display: flex; flex-wrap: wrap; gap: 24px;
        font-variant-numeric: tabular-nums;
    }}
    .ecng-ticker .item {{ color: #d9e2ec; font-size: 0.8rem; font-weight: 500; white-space: nowrap; }}
    .ecng-ticker .item b {{ color: {GOLD}; font-weight: 700; margin-left: 6px; }}

    .ecng-timestamp-pill {{
        display: inline-flex; align-items: center; gap: 6px;
        background: #eef2f8; border: 1px solid #dbe3ee; border-radius: 999px;
        padding: 4px 12px; font-size: 0.75rem; color: #5b6472; margin: 12px 0 16px 0;
    }}

    /* Home nav buttons — shared base: big, gradient fill, shadow, tappable, pill-shaped */
    div[data-testid="stVerticalBlock"] .stButton > button {{
        width: 100%; padding: 24px 20px; font-size: 1.08rem; font-weight: 700;
        font-family: 'Poppins', sans-serif;
        border-radius: 40px; border: none; text-align: left; color: #fff;
        margin-bottom: 10px; box-shadow: 0 6px 14px rgba(0,0,0,0.18);
        transition: transform 0.08s ease, box-shadow 0.08s ease;
    }}
    div[data-testid="stVerticalBlock"] .stButton > button:hover {{
        transform: translateY(-1px); box-shadow: 0 8px 18px rgba(0,0,0,0.22); color: #fff;
    }}
    div[data-testid="stVerticalBlock"] .stButton > button:active {{ transform: translateY(0px); }}

    /* Per-button color identity, matched to the same meanings used elsewhere in the app */
    .st-key-nav_outstanding_wrap button {{ background: linear-gradient(135deg, #003d8f 0%, {NAVY} 100%); }}
    .st-key-nav_transacted_wrap button {{ background: linear-gradient(135deg, #ffdb4d 0%, {GOLD} 100%); color: {NAVY} !important; }}
    .st-key-nav_pricing_wrap button {{ background: linear-gradient(135deg, #5ecb4a 0%, {GREEN} 100%); }}

    /* Back button — deliberately small and high-contrast, distinct from the big primary nav buttons */
    .st-key-back_home_wrap button {{
        width: auto !important; padding: 7px 18px !important; font-size: 0.82rem !important;
        font-family: inherit !important; border-radius: 999px !important;
        background: #fff !important; color: {NAVY} !important; border: 2px solid {NAVY} !important;
        box-shadow: none !important; margin-bottom: 14px !important;
    }}
    .st-key-back_home_wrap button:hover {{
        background: {NAVY} !important; color: #fff !important;
    }}

    .nav-caption {{ font-size: 0.78rem; color: #6b7280; margin: -4px 0 16px 4px; }}

    .ecng-card {{
        border: 1px solid #eceff2; border-left: var(--accent-width, 4px) solid var(--accent, {GRAY});
        border-radius: 14px; padding: 12px 16px; margin-bottom: 10px;
        background: #fff url('data:image/png;base64,{WATERMARK_B64}') no-repeat right 8px bottom 8px / 90px auto;
        box-shadow: 0 3px 10px rgba(0,0,0,0.06);
    }}
    .ecng-card .client {{ font-weight: 700; font-size: 0.98rem; color: {DARK_GRAY}; }}
    .ecng-card .meta {{ font-size: 0.8rem; color: #6b7280; margin-top: 1px; }}
    .ecng-card .row {{
        display: flex; justify-content: space-between; align-items: center;
        margin-top: 6px; font-size: 0.85rem; font-variant-numeric: tabular-nums;
    }}
    .ecng-card .row + .row {{
        border-top: 1px solid #f0f2f5; padding-top: 7px; margin-top: 8px;
    }}
    .ecng-label {{
        display: block; font-size: 0.65rem; text-transform: uppercase;
        letter-spacing: 0.07em; color: #9aa1ab; font-weight: 500; margin-bottom: 1px;
    }}
    .ecng-value {{ font-size: 0.92rem; color: {DARK_GRAY}; font-weight: 600; }}

    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div {{
        border-color: {NAVY} !important; border-radius: 14px !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.06);
    }}
    div[data-testid="stSelectbox"] label {{ color: {NAVY} !important; font-weight: 600; font-size: 0.8rem !important; }}

    .ecng-avatar {{
        display: inline-flex; align-items: center; justify-content: center;
        width: 20px; height: 20px; border-radius: 50%;
        background: linear-gradient(135deg, #003d8f 0%, {NAVY} 100%); color: {GOLD};
        font-size: 0.6rem; font-weight: 700; margin-right: 5px; vertical-align: middle;
    }}

    .ecng-price-card {{
        border: 1px solid #eceff2; border-radius: 12px; padding: 10px 14px;
        margin-bottom: 6px; background: #fafbfc; box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }}
    .ecng-price-card .hub {{ font-size: 0.78rem; color: #6b7280; }}
    .ecng-price-card .price {{ font-size: 1.05rem; font-weight: 700; color: {NAVY}; }}
    .ecng-period-badge {{
        background-color: {NAVY} !important; display: flex; align-items: center;
        justify-content: center; min-width: 92px;
    }}
    .ecng-period-badge .price {{ color: #fff; font-size: 0.82rem; font-weight: 700; }}
    .ecng-curve-row {{ display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 8px; }}
    .ecng-curve-row .ecng-price-card {{ flex: 1 1 90px; margin-bottom: 0; }}
    .ecng-section-pill {{
        background-color: {NAVY}; color: #fff; font-weight: 700; font-size: 0.82rem;
        font-family: 'Poppins', sans-serif;
        padding: 6px 14px; border-radius: 999px; display: inline-block; margin-bottom: 8px;
        box-shadow: 0 3px 8px rgba(0,47,108,0.25);
    }}
    .ecng-page-title-divider {{
        border: none; height: 2px; background-color: {GOLD}; opacity: 0.8;
        margin: 4px 0 16px 0; width: 60px;
    }}

    .ecng-empty {{
        background-color: #f8f9fb; border: 1px solid #e6e9ee; border-left: 4px solid {GRAY};
        border-radius: 12px; padding: 16px 16px; color: {DARK_GRAY}; font-size: 0.85rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }}
    .ecng-stale {{
        background-color: #fff8e1; border: 1px solid {GOLD}; border-radius: 6px;
        padding: 8px 12px; font-size: 0.78rem; color: {DARK_GRAY}; margin-bottom: 10px;
    }}
</style>
""", unsafe_allow_html=True)


def render_header(with_ticker=False):
    st.markdown(f"""
    <div class="ecng-header">
        <div class="ecng-header-row">
            <div class="ecng-logo-badge"><img src="data:image/png;base64,{LOGO_B64}" alt="ECNG logo" /></div>
            <div>
                <h1>ECNG Energy Group</h1>
                <p>Outstanding terms &middot; pricing &middot; mobile snapshot</p>
            </div>
        </div>
    </div>
    <div class="ecng-header-divider{'' if with_ticker else ' rounded'}"></div>
    """, unsafe_allow_html=True)


def render_ticker(snapshot):
    gas_strip = next((s for s in snapshot.get("pricing", []) if s["name"] == "Gas Strip"), None)
    if not gas_strip or not gas_strip["rows"]:
        return
    # forward term (the one after prompt), since ECNG buys forward — falls
    # back to the first row if there's only one available
    rows = gas_strip["rows"]
    nearest = rows[1] if len(rows) > 1 else rows[0]
    items = "".join(
        f'<span class="item">{h}<b>{nearest["prices"].get(h, "—")}</b></span>'
        for h in gas_strip["hubs"]
    )
    st.markdown(
        f'<div class="ecng-ticker">'
        f'<span class="item" style="color:#8fa5c7;">{nearest["period_label"].upper()}</span>'
        f'{items}</div>',
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------- data loading
@st.cache_data(ttl=300)
def load_snapshot():
    try:
        share_url = st.secrets.get("ONEDRIVE_SHARE_URL", "").strip()
    except Exception:
        share_url = ""
    if share_url:
        try:
            xlsx_bytes = onedrive_source.fetch_excel_bytes(share_url)
            return excel_parser.parse_workbook_bytes(xlsx_bytes)
        except Exception as e:
            st.session_state["_onedrive_error"] = str(e)

    data_path = Path(__file__).parent / "data.json"
    if not data_path.exists():
        return None
    return json.loads(data_path.read_text())


# ---------------------------------------------------------------- helpers
def rep_initials(name):
    if not name:
        return "?"
    parts = str(name).strip().split()
    return (parts[0][:2] if len(parts) == 1 else parts[0][0] + parts[-1][0]).upper()


def fmt_date(v):
    if not v:
        return "—"
    try:
        return datetime.strptime(v, "%Y-%m-%d").strftime("%b %-d, %Y")
    except ValueError:
        return v


def term_str(start, end):
    def short(v):
        try:
            return datetime.strptime(v, "%Y-%m-%d").strftime("%b%y")
        except (ValueError, TypeError):
            return None
    s, e = short(start), short(end)
    if s and e:
        return f"{s}-{e}"
    return s or e or ""


def price_gap(target, market):
    if target is None or market is None:
        return None
    return round(abs(target - market), 2)


def urgency_style(status, target, market):
    """Returns (color, border_width). In the Money is always green and a
    touch thicker, regardless of price gap — it's a status signal, not a
    gap signal. Everything else is colored by how far market has moved
    from target: yellow up to 20c, blue beyond that."""
    if status == "In the Money":
        return GREEN, "6px"
    gap = price_gap(target, market)
    if gap is None:
        return GRAY, "4px"
    if gap > 0.10:
        return NAVY, "4px"   # "blue"
    return GOLD, "4px"       # "yellow" — covers 0 up to 10c


def sort_key(s):
    gap = price_gap(s.get("target"), s.get("market"))
    return gap if gap is not None else float("inf")


def price_str(v):
    return f"${v:,.2f}" if v is not None else "—"


def volume_str(v, product):
    if v is None:
        return "—"
    unit = "GJ" if product and product.strip().lower() == "gas" else "kW"
    return f"{v:,.0f} {unit}"


def delta_str(target, market):
    if target is None or market is None:
        return "—"
    gap = target - market
    if gap > 0:
        return f'<span style="color:{GREEN};font-weight:700;">▼ ${gap:,.2f}</span>'
    if gap < 0:
        return f'<span style="color:{RED};font-weight:700;">▲ ${abs(gap):,.2f}</span>'
    return "$0.00"


def render_deal_card(s):
    is_transacted = s["status"] == "Transacted"
    status_color = STATUS_COLORS.get(s["status"], GRAY)
    is_ab = (s.get("product") or "").strip().lower().startswith("ab")
    volume_row = "" if is_ab else (
        '<div class="row"><div>'
        '<span class="ecng-label">Volume</span>'
        f'<span class="ecng-value">{volume_str(s.get("volume"), s.get("product"))}</span>'
        '</div></div>'
    )

    if is_transacted:
        color = GOLD
        savings = s.get("savings_vs_target")
        if savings is None:
            savings_html = f'<span style="color:{GRAY};">—</span>'
        elif savings < 0:
            savings_html = f'<span style="color:{GREEN};font-weight:700;">▼ Transacted ${abs(savings):,.2f} Below Target</span>'
        elif savings > 0:
            savings_html = f'<span style="color:{RED};font-weight:700;">▲ Transacted ${savings:,.2f} Above Target</span>'
        else:
            savings_html = '<span style="font-weight:700;">Transacted On Target</span>'
        notes = s.get("notes")
        notes_row = (
            f'<div class="row"><div><span class="ecng-label">Notes</span>'
            f'<span class="ecng-value">{notes}</span></div></div>'
        ) if notes else ""
        return (
            f'<div class="ecng-card" style="--accent:{color};">'
            f'<div class="client">{s["client"]}</div>'
            f'<div class="meta"><span class="ecng-avatar">{rep_initials(s["rep"])}</span>'
            f'{s["rep"]} &middot; {s["product"]} &middot; {s.get("delivery_type") or ""}</div>'
            f'{volume_row}'
            '<div class="row" style="align-items:flex-start;"><div>'
            '<span class="ecng-label">Target</span>'
            f'<span class="ecng-value">{price_str(s["target"])}</span>'
            '</div><div style="text-align:right;">'
            '<span class="ecng-label">Transacted Price</span>'
            f'<div class="ecng-value">{price_str(s.get("transacted_price"))}</div>'
            f'<div>{savings_html}</div>'
            '</div></div>'
            '<div class="row" style="align-items:flex-start;"><div>'
            '<span class="ecng-label">Date Transacted</span>'
            f'<span class="ecng-value">{fmt_date(s.get("transacted_date"))}</span>'
            '</div></div>'
            f'{notes_row}'
            '</div>'
        )
    else:
        color, width = urgency_style(s["status"], s.get("target"), s.get("market"))
        days = s.get("days_to_expiry")
        days_str = f"{days}d left" if days is not None else "—"
        return (
            f'<div class="ecng-card" style="--accent:{color}; --accent-width:{width};">'
            f'<div class="client">{s["client"]}</div>'
            f'<div class="meta"><span class="ecng-avatar">{rep_initials(s["rep"])}</span>'
            f'{s["rep"]} &middot; {s["product"]} &middot; {s.get("delivery_type") or ""}</div>'
            f'{volume_row}'
            '<div class="row" style="align-items:flex-start;"><div>'
            '<span class="ecng-label">Target</span>'
            f'<span class="ecng-value">{price_str(s["target"])}</span>'
            '</div><div style="text-align:right;">'
            f'<div style="font-weight:700; color:{DARK_GRAY}; margin-bottom:4px;">{term_str(s["start_date"], s["end_date"])}</div>'
            '<span class="ecng-label">Market</span>'
            f'<div class="ecng-value">{price_str(s["market"])}</div>'
            f'<div>{delta_str(s["target"], s["market"])}</div>'
            '</div></div>'
            '<div class="row" style="align-items:flex-start;"><div>'
            '<span class="ecng-label">Expiry</span>'
            f'<span class="ecng-value">{fmt_date(s["expiry_date"])}</span>'
            '</div>'
            f'<div style="color:{status_color};font-weight:700;">{days_str}</div>'
            '</div></div>'
        )


def render_freshness_banner(snapshot):
    gen_at = snapshot.get("generated_at")
    if not gen_at:
        return
    try:
        gen_dt = datetime.fromisoformat(gen_at)
        age_hours = (datetime.now() - gen_dt).total_seconds() / 3600
        label = gen_dt.strftime("%b %-d, %Y at %-I:%M %p")
        if age_hours > 20:
            st.markdown(
                f'<div class="ecng-stale">⚠️ This snapshot is from {label} — it may be a day or more old.</div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(f'<div class="ecng-timestamp-pill">🕐 Snapshot as of {label}</div>', unsafe_allow_html=True)
    except ValueError:
        st.markdown(f'<div class="ecng-timestamp-pill">🕐 Snapshot as of {gen_at}</div>', unsafe_allow_html=True)


# ---------------------------------------------------------------- pages
def go(page):
    st.session_state.page = page


def page_home(snapshot):
    render_header(with_ticker=True)
    render_ticker(snapshot)
    if st.session_state.get("_onedrive_error"):
        st.markdown(
            f'<div class="ecng-stale">⚠️ Could not reach OneDrive — showing the last saved '
            f'snapshot instead.<br><span style="font-size:0.7rem;">{st.session_state["_onedrive_error"]}</span></div>',
            unsafe_allow_html=True,
        )
    render_freshness_banner(snapshot)

    swaps = snapshot.get("swaps", [])
    outstanding_count = sum(1 for s in swaps if s["status"] in ("Active", "In the Money", "Expired"))
    transacted_count = sum(1 for s in swaps if s["status"] == "Transacted")
    pricing_count = len(snapshot.get("pricing", []))

    st.write("")
    with st.container(key="nav_outstanding_wrap"):
        st.button(f"📋  Outstanding  ·  {outstanding_count}", key="nav_outstanding",
                  on_click=go, args=("outstanding",), use_container_width=True)
    st.markdown('<div class="nav-caption">Shows Active + In the Money together, filterable individually</div>', unsafe_allow_html=True)

    with st.container(key="nav_transacted_wrap"):
        st.button(f"💰  Transacted  ·  {transacted_count}", key="nav_transacted",
                  on_click=go, args=("transacted",), use_container_width=True)
    st.markdown('<div class="nav-caption">Deals that have been executed</div>', unsafe_allow_html=True)

    with st.container(key="nav_pricing_wrap"):
        st.button(f"🏷️  Pricing  ·  {pricing_count} curve" + ("s" if pricing_count != 1 else ""),
                  key="nav_pricing", on_click=go, args=("pricing",), use_container_width=True)
    st.markdown('<div class="nav-caption">Current gas and power pricing curves</div>', unsafe_allow_html=True)


def page_outstanding(snapshot):
    with st.container(key="back_home_wrap"):
        st.button("← Home", on_click=go, args=("home",))
    render_header()

    swaps = snapshot.get("swaps", [])
    reps = sorted({s["rep"] for s in swaps if s.get("rep")})
    products = sorted({s["product"] for s in swaps if s.get("product")})

    sel_rep = st.selectbox("My book", ["All reps"] + reps)
    col1, col2 = st.columns(2)
    with col1:
        sel_product = st.selectbox("Product", ["All products"] + products)
    with col2:
        view = st.selectbox("View", ["Active + In the Money", "Active", "In the Money", "Expired"])

    filtered = [s for s in swaps if sel_rep == "All reps" or s["rep"] == sel_rep]
    filtered = [s for s in filtered if sel_product == "All products" or s["product"] == sel_product]
    if view == "Active + In the Money":
        filtered = [s for s in filtered if s["status"] in ("Active", "In the Money")]
    else:
        filtered = [s for s in filtered if s["status"] == view]
    filtered.sort(key=sort_key)

    st.markdown(f"**{len(filtered)}** {view.lower()} deal(s)")

    if not filtered:
        st.markdown('<div class="ecng-empty">📭 &nbsp;Nothing here right now.</div>', unsafe_allow_html=True)
    else:
        for s in filtered:
            st.markdown(render_deal_card(s), unsafe_allow_html=True)


def page_transacted(snapshot):
    with st.container(key="back_home_wrap"):
        st.button("← Home", on_click=go, args=("home",))
    render_header()

    swaps = [s for s in snapshot.get("swaps", []) if s["status"] == "Transacted"]
    reps = sorted({s["rep"] for s in swaps if s.get("rep")})
    products = sorted({s["product"] for s in swaps if s.get("product")})

    sel_rep = st.selectbox("Rep", ["All reps"] + reps)
    sel_product = st.selectbox("Product", ["All products"] + products)

    filtered = [s for s in swaps if sel_rep == "All reps" or s["rep"] == sel_rep]
    filtered = [s for s in filtered if sel_product == "All products" or s["product"] == sel_product]
    filtered.sort(key=lambda s: s.get("transacted_date") or "", reverse=True)

    st.markdown(f"**{len(filtered)}** transacted deal(s)")

    if not filtered:
        st.markdown('<div class="ecng-empty">📭 &nbsp;No transacted deals match the current filters.</div>', unsafe_allow_html=True)
    else:
        for s in filtered:
            st.markdown(render_deal_card(s), unsafe_allow_html=True)


def page_pricing(snapshot):
    with st.container(key="back_home_wrap"):
        st.button("← Home", on_click=go, args=("home",))
    render_header()
    st.subheader("🏷️ Current Pricing")
    st.markdown('<div class="ecng-page-title-divider"></div>', unsafe_allow_html=True)

    for section in snapshot.get("pricing", []):
        st.markdown(f'<div class="ecng-section-pill">{section["name"]}</div>', unsafe_allow_html=True)
        rows = section.get("rows", [])
        if not rows:
            continue
        for r in rows:
            hub_bubbles = "".join(
                f'<div class="ecng-price-card"><div class="hub">{h}</div>'
                f'<div class="price">{r["prices"].get(h, "—").replace("$", "&#36;")}</div></div>'
                for h in section["hubs"]
            )
            st.markdown(
                f'<div class="ecng-curve-row">'
                f'<div class="ecng-price-card ecng-period-badge"><div class="price">{r["period_label"]}</div></div>'
                f'{hub_bubbles}'
                f'</div>',
                unsafe_allow_html=True,
            )


# ---------------------------------------------------------------- main
if "page" not in st.session_state:
    st.session_state.page = "home"

snapshot = load_snapshot()

if not snapshot:
    render_header()
    st.markdown(
        '<div class="ecng-empty">📭 &nbsp;No data yet — waiting on the first snapshot.</div>',
        unsafe_allow_html=True,
    )
    st.stop()

page = st.session_state.page
if page == "home":
    page_home(snapshot)
elif page == "outstanding":
    page_outstanding(snapshot)
elif page == "transacted":
    page_transacted(snapshot)
elif page == "pricing":
    page_pricing(snapshot)
else:
    st.session_state.page = "home"
    st.rerun()
