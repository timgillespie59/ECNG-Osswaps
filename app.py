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

# ---------------------------------------------------------------- styles
st.markdown(f"""
<style>
    .block-container {{ padding-top: 1rem; padding-bottom: 3rem; max-width: 480px; }}
    .ecng-header {{
        position: relative; background-color: {NAVY}; padding: 16px 20px 14px 20px;
        border-radius: 8px; margin-bottom: 6px; overflow: hidden;
    }}
    .ecng-header h1 {{ color: #fff; margin: 0; font-size: 1.35rem; position: relative; z-index: 1; }}
    .ecng-header p {{ color: #b9c9e2; margin: 2px 0 0 0; font-size: 0.8rem; position: relative; z-index: 1; }}
    .ecng-header svg {{ position: absolute; bottom: 0; left: 0; right: 0; width: 100%; height: 36px; opacity: 0.5; }}

    .ecng-ticker {{
        background-color: {NAVY}; border-top: 1px solid rgba(255,255,255,0.15);
        border-radius: 0 0 6px 6px; margin-top: -6px; margin-bottom: 18px;
        padding: 8px 20px; display: flex; flex-wrap: wrap; gap: 24px;
        font-variant-numeric: tabular-nums;
    }}
    .ecng-ticker .item {{ color: #d9e2ec; font-size: 0.8rem; font-weight: 500; white-space: nowrap; }}
    .ecng-ticker .item b {{ color: {GOLD}; font-weight: 700; margin-left: 6px; }}

    /* Home nav buttons — big, brand-colored, tappable */
    div[data-testid="stVerticalBlock"] .stButton > button {{
        width: 100%; padding: 22px 18px; font-size: 1.05rem; font-weight: 700;
        border-radius: 12px; border: none; text-align: left;
        background-color: {NAVY}; color: #fff; margin-bottom: 12px;
    }}
    div[data-testid="stVerticalBlock"] .stButton > button:hover {{
        background-color: #003d8f; color: #fff;
    }}
    .nav-caption {{ font-size: 0.78rem; color: #6b7280; margin: -8px 0 16px 4px; }}

    .ecng-card {{
        border: 1px solid #e6e6e6; border-left: 4px solid var(--accent, {GRAY});
        border-radius: 8px; padding: 10px 14px; margin-bottom: 8px; background: #fff;
    }}
    .ecng-card .client {{ font-weight: 700; font-size: 0.98rem; color: {DARK_GRAY}; }}
    .ecng-card .meta {{ font-size: 0.8rem; color: #6b7280; margin-top: 1px; }}
    .ecng-card .row {{
        display: flex; justify-content: space-between; align-items: center;
        margin-top: 6px; font-size: 0.85rem; font-variant-numeric: tabular-nums;
    }}
    .ecng-label {{
        display: block; font-size: 0.65rem; text-transform: uppercase;
        letter-spacing: 0.04em; color: #9aa1ab; font-weight: 600; margin-bottom: 1px;
    }}
    .ecng-value {{ font-size: 0.92rem; color: {DARK_GRAY}; font-weight: 600; }}

    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div {{
        border-color: {NAVY} !important; border-radius: 8px !important;
    }}
    div[data-testid="stSelectbox"] label {{ color: {NAVY} !important; font-weight: 600; font-size: 0.8rem !important; }}

    .ecng-avatar {{
        display: inline-flex; align-items: center; justify-content: center;
        width: 20px; height: 20px; border-radius: 50%; background-color: {NAVY}; color: {GOLD};
        font-size: 0.6rem; font-weight: 700; margin-right: 5px; vertical-align: middle;
    }}

    .ecng-price-card {{
        border: 1px solid #e6e6e6; border-radius: 8px; padding: 10px 14px;
        margin-bottom: 6px; background: #fafbfc;
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
        padding: 6px 14px; border-radius: 999px; display: inline-block; margin-bottom: 8px;
    }}

    .ecng-empty {{
        background-color: #f5f6f8; border: 1px dashed #cdd3da; border-left: 4px solid {GRAY};
        border-radius: 6px; padding: 12px 14px; color: {DARK_GRAY}; font-size: 0.85rem;
    }}
    .ecng-stale {{
        background-color: #fff8e1; border: 1px solid {GOLD}; border-radius: 6px;
        padding: 8px 12px; font-size: 0.78rem; color: {DARK_GRAY}; margin-bottom: 10px;
    }}
</style>
""", unsafe_allow_html=True)


def render_header():
    st.markdown(f"""
    <div class="ecng-header">
        <h1>ECNG Energy Group</h1>
        <p>Outstanding terms &middot; pricing &middot; mobile snapshot</p>
        <svg viewBox="0 0 800 60" preserveAspectRatio="none" xmlns="http://www.w3.org/2000/svg">
            <polyline points="0,45 80,30 160,38 240,15 320,25 400,10 480,22 560,8 640,18 720,5 800,14"
                      fill="none" stroke="{GOLD}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
    </div>
    """, unsafe_allow_html=True)


def render_ticker(snapshot):
    gas_strip = next((s for s in snapshot.get("pricing", []) if s["name"] == "Gas Strip"), None)
    if not gas_strip or not gas_strip["rows"]:
        return
    nearest = gas_strip["rows"][0]
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


def urgency_color(target, market):
    gap = price_gap(target, market)
    if gap is None:
        return GRAY
    if gap <= 0.05:
        return GOLD
    if gap > 0.20:
        return GRAY
    return NAVY


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
            savings_html = f'<span style="color:{GREEN};font-weight:700;">▼ ${abs(savings):,.2f} saved</span>'
        elif savings > 0:
            savings_html = f'<span style="color:{RED};font-weight:700;">▲ ${savings:,.2f} over target</span>'
        else:
            savings_html = '<span style="font-weight:700;">On target</span>'
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
        color = urgency_color(s.get("target"), s.get("market"))
        days = s.get("days_to_expiry")
        days_str = f"{days}d left" if days is not None else "—"
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
            st.caption(f"Snapshot as of {label}")
    except ValueError:
        st.caption(f"Snapshot as of {gen_at}")


# ---------------------------------------------------------------- pages
def go(page):
    st.session_state.page = page


def page_home(snapshot):
    render_header()
    render_ticker(snapshot)
    if st.session_state.get("_onedrive_error"):
        st.markdown(
            f'<div class="ecng-stale">⚠️ Could not reach OneDrive — showing the last saved '
            f'snapshot instead.<br><span style="font-size:0.7rem;">{st.session_state["_onedrive_error"]}</span></div>',
            unsafe_allow_html=True,
        )
    render_freshness_banner(snapshot)

    st.write("")
    st.button("📋  Outstanding", key="nav_outstanding", on_click=go, args=("outstanding",), use_container_width=True)
    st.markdown('<div class="nav-caption">Active, In the Money, and Expired terms</div>', unsafe_allow_html=True)

    st.button("💰  Transacted", key="nav_transacted", on_click=go, args=("transacted",), use_container_width=True)
    st.markdown('<div class="nav-caption">Deals that have been executed</div>', unsafe_allow_html=True)

    st.button("🏷️  Pricing", key="nav_pricing", on_click=go, args=("pricing",), use_container_width=True)
    st.markdown('<div class="nav-caption">Current gas and power pricing curves</div>', unsafe_allow_html=True)


def page_outstanding(snapshot):
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
        view = st.selectbox("View", ["Active", "In the Money", "Expired"])

    filtered = [s for s in swaps if sel_rep == "All reps" or s["rep"] == sel_rep]
    filtered = [s for s in filtered if sel_product == "All products" or s["product"] == sel_product]
    filtered = [s for s in filtered if s["status"] == view]
    filtered.sort(key=sort_key)

    st.markdown(f"**{len(filtered)}** {view.lower()} deal(s)")

    if not filtered:
        st.markdown('<div class="ecng-empty">Nothing here right now.</div>', unsafe_allow_html=True)
    else:
        for s in filtered:
            st.markdown(render_deal_card(s), unsafe_allow_html=True)


def page_transacted(snapshot):
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
        st.markdown('<div class="ecng-empty">No transacted deals match the current filters.</div>', unsafe_allow_html=True)
    else:
        for s in filtered:
            st.markdown(render_deal_card(s), unsafe_allow_html=True)


def page_pricing(snapshot):
    st.button("← Home", on_click=go, args=("home",))
    render_header()
    st.subheader("🏷️ Current Pricing")

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
        '<div class="ecng-empty">No data yet — waiting on the first snapshot.</div>',
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
