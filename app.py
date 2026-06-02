import pandas as pd
import plotly.express as px
import streamlit as st
from pathlib import Path

st.set_page_config(page_title='RealKPI vs NoteUserKPI', layout='wide')

DATA_MASTER = Path('MASTER DATA.xlsx')
DATA_FORMATTING = Path('FORMATTING.xlsx')
ID_COLS = ['Region', 'KPI INDICES', 'PIC', 'Kategori']

COLOR_BG = '#F2F3F5'
COLOR_CARD = '#FFFFFF'
COLOR_PRIMARY = '#862880'
COLOR_ACCENT = '#EE6825'
COLOR_TEXT = '#1D1D1F'
COLOR_MUTED = '#5F6368'
COLOR_BORDER = '#D7D9DE'

REAL_COLOR = COLOR_PRIMARY
NOTE_COLOR = COLOR_ACCENT
INC_COLOR = COLOR_ACCENT
BG_CARD = COLOR_CARD


def inject_styles() -> None:
    st.markdown(
        f"""
        <style>
        :root {{
            --app-text: {COLOR_TEXT};
            --app-subtext: {COLOR_MUTED};
            --app-card-bg: {BG_CARD};
            --app-card-border: {COLOR_BORDER};
            --app-bg: {COLOR_BG};
            --sidebar-bg: {COLOR_CARD};
            --tab-text: {COLOR_TEXT};
            --app-primary: {COLOR_PRIMARY};
            --app-accent: {COLOR_ACCENT};
        }}
        .stApp {{
            background: var(--app-bg);
            color: var(--app-text);
        }}
        .stMarkdown, .stText, p, h1, h2, h3, h4, h5, h6,
        label, span, div[data-testid='stMarkdownContainer'] {{
            color: var(--app-text);
        }}
        .stSidebar h1, .stSidebar h2, .stSidebar h3, .stSidebar h4,
        .stSidebar label, .stSidebar p, .stSidebar span, .stSidebar div {{
            color: var(--app-text) !important;
        }}
        .hero {{
            padding: 18px 22px;
            border-radius: 8px;
            background: linear-gradient(90deg, var(--app-primary) 0%, #A83C9F 62%, var(--app-accent) 100%);
            color: #ffffff;
            margin-bottom: 14px;
            border: 1px solid rgba(134, 40, 128, 0.25);
        }}
        .hero h2 {{ margin: 0; font-size: 28px; letter-spacing: 0.3px; }}
        .hero h2, .hero p {{ color: #ffffff !important; }}
        .hero p {{ margin: 6px 0 0; opacity: 0.92; }}
        .kpi-card {{
            background: var(--app-card-bg);
            border: 1px solid var(--app-card-border);
            border-radius: 8px;
            padding: 10px 12px;
            min-height: 86px;
            box-shadow: 0 1px 2px rgba(29, 29, 31, 0.06);
        }}
        .kpi-title {{ font-size: 12px; color: var(--app-subtext); margin-bottom: 4px; }}
        .kpi-value {{ font-size: 27px; font-weight: 700; color: var(--app-text); margin: 0; }}
        .kpi-hint {{ font-size: 11px; color: var(--app-subtext); }}
        div[data-testid='stSidebar'] {{
            background: var(--sidebar-bg);
            border-right: 1px solid var(--app-card-border);
        }}
        button[data-baseweb='tab'] {{
            color: var(--tab-text) !important;
        }}
        button[data-baseweb='tab'][aria-selected='true'] {{
            color: var(--app-primary) !important;
        }}
        div[data-testid='stPopoverBody'],
        div[data-baseweb='popover'] {{
            background: var(--app-card-bg) !important;
            color: var(--app-text) !important;
        }}
        div[data-baseweb='select'] > div,
        div[data-baseweb='input'] > div {{
            background: #ffffff !important;
            border-color: var(--app-card-border) !important;
            color: var(--app-text) !important;
        }}
        div[data-baseweb='tag'] {{
            background-color: var(--app-primary) !important;
            color: #ffffff !important;
        }}
        div[data-baseweb='tag'],
        div[data-baseweb='tag'] *,
        div[data-baseweb='tag'] span,
        div[data-baseweb='tag'] svg,
        div[data-baseweb='tag'] button {{
            color: #ffffff !important;
            fill: #ffffff !important;
        }}
        h3, .stSubheader {{
            color: var(--app-text) !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data
def load_data(master_path: Path, formatting_path: Path) -> pd.DataFrame:
    master_df = pd.read_excel(master_path)
    formatting_df = pd.read_excel(formatting_path)

    master_df.columns = [str(c).strip() for c in master_df.columns]
    formatting_df.columns = [str(c).strip() for c in formatting_df.columns]

    formatting_df = formatting_df.rename(
        columns={
            'INDICES': 'KPI Indices',
            'FORMAT': 'Measurement Type',
        }
    )
    format_map = (
        formatting_df[['KPI Indices', 'Measurement Type', 'KPI Category', 'Direction']]
        .dropna(subset=['KPI Indices', 'Measurement Type'])
        .assign(
            **{
                'KPI Indices': lambda d: d['KPI Indices'].astype(str).str.strip(),
                'Measurement Type': lambda d: d['Measurement Type'].astype(str).str.strip().str.lower(),
                'KPI Category': lambda d: d['KPI Category'].fillna('Uncategorized').astype(str).str.strip(),
                'Direction': lambda d: d['Direction'].fillna('Higher Better').astype(str).str.strip(),
            }
        )
    )

    df = master_df.merge(format_map, left_on='KPI INDICES', right_on='KPI Indices', how='left')
    df['Measurement Type'] = df['Measurement Type'].fillna('absolute number')
    df['KPI Category'] = df['KPI Category'].fillna('Uncategorized')
    df['Direction'] = df['Direction'].fillna('Higher Better')
    df['MetricType'] = df['Measurement Type'].map(
        {
            'percentage': 'percentage',
            'absolute number': 'absolute',
        }
    ).fillna('absolute')
    df['FormatLabel'] = df['MetricType'].map({'percentage': 'pct_format', 'absolute': 'abs_format'})
    df['MetricFormat'] = df['Measurement Type']

    on_time = pd.to_numeric(df['ON TIME'], errors='coerce')
    over_time = pd.to_numeric(df['OVER TIME'], errors='coerce')
    current_over_time = pd.to_numeric(df['Current OVER TIME'], errors='coerce')
    total_os = pd.to_numeric(df['TOTAL OS'], errors='coerce')
    current_total_os = pd.to_numeric(df['Current TOTAL OS'], errors='coerce')

    denom_real = on_time + over_time
    denom_note = on_time + current_over_time

    real_pct = on_time.div(denom_real.where(denom_real.ne(0)))
    note_pct = on_time.div(denom_note.where(denom_note.ne(0)))

    df['ON TIME'] = on_time
    df['OVER TIME'] = over_time
    df['Current OVER TIME'] = current_over_time
    df['TOTAL OS'] = total_os
    df['Current TOTAL OS'] = current_total_os
    df['RealDenom'] = denom_real
    df['NoteDenom'] = denom_note

    df['RealKPI'] = total_os
    df['NoteUserKPI'] = current_total_os
    pct_mask = df['MetricType'].eq('percentage')
    df.loc[pct_mask, 'RealKPI'] = real_pct[pct_mask]
    df.loc[pct_mask, 'NoteUserKPI'] = note_pct[pct_mask]

    df['Increase_abs'] = df['NoteUserKPI'] - df['RealKPI']
    df['Increase_pct_vs_real'] = df['Increase_abs'].div(df['RealKPI'].where(df['RealKPI'].ne(0))) * 100
    return df


def aggregate_metrics(df: pd.DataFrame, group_cols: list[str], metric_type: str) -> pd.DataFrame:
    grouped = df.groupby(group_cols, as_index=False) if group_cols else None

    if metric_type == 'percentage':
        if grouped is None:
            agg = pd.DataFrame(
                {
                    'ON TIME': [df['ON TIME'].sum()],
                    'RealDenom': [df['RealDenom'].sum()],
                    'NoteDenom': [df['NoteDenom'].sum()],
                }
            )
        else:
            agg = grouped[['ON TIME', 'RealDenom', 'NoteDenom']].sum()

        agg['RealKPI'] = agg['ON TIME'].div(agg['RealDenom'].where(agg['RealDenom'].ne(0)))
        agg['NoteUserKPI'] = agg['ON TIME'].div(agg['NoteDenom'].where(agg['NoteDenom'].ne(0)))
    else:
        if grouped is None:
            agg = pd.DataFrame(
                {
                    'RealKPI': [df['TOTAL OS'].sum()],
                    'NoteUserKPI': [df['Current TOTAL OS'].sum()],
                }
            )
        else:
            agg = grouped.agg(RealKPI=('TOTAL OS', 'sum'), NoteUserKPI=('Current TOTAL OS', 'sum'))

    agg['Increase_abs'] = agg['NoteUserKPI'] - agg['RealKPI']
    agg['Increase_pct_vs_real'] = agg['Increase_abs'].div(agg['RealKPI'].where(agg['RealKPI'].ne(0))) * 100
    return agg


def direction_label(direction: str) -> str:
    normalized = str(direction).strip().lower()
    if normalized == 'higher better':
        return 'Target Naik'
    if normalized == 'lower better':
        return 'Target Turun'
    return str(direction)


def category_label(category: str) -> str:
    normalized = str(category).strip().lower()
    if normalized == 'main kpi':
        return 'KPI Utama'
    if normalized == 'support kpi':
        return 'KPI Pendukung'
    return str(category)


def metric_label(metric_type: str) -> str:
    if metric_type == 'percentage':
        return 'Persentase'
    if metric_type == 'absolute':
        return 'Angka Absolut'
    return str(metric_type)


def _filter_popover(label: str, options: list[str], key: str) -> list[str]:
    selected_key = f'{key}_selected'
    widget_key = f'{key}_ms'
    if selected_key not in st.session_state:
        st.session_state[selected_key] = options.copy()

    selected = [v for v in st.session_state[selected_key] if v in options]
    widget_selected = st.session_state.get(widget_key, selected)
    widget_selected = [v for v in widget_selected if v in options]

    # Keep both states aligned so summary count and multiselect always match.
    st.session_state[selected_key] = widget_selected
    st.session_state[widget_key] = widget_selected
    selected = widget_selected
    summary = f'{label} ({len(selected)}/{len(options)})'

    with st.sidebar.popover(summary, use_container_width=True):
        current_default = st.session_state.get(widget_key, st.session_state[selected_key])
        picked = st.multiselect(
            f'Cari {label}',
            options=options,
            default=current_default,
            key=widget_key,
            placeholder=f'Ketik untuk search {label}',
        )
        st.session_state[selected_key] = picked

    return st.session_state[selected_key]


def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.header('Filter Dashboard')

    regions = sorted(df['Region'].dropna().unique().tolist())
    indices = sorted(df['KPI INDICES'].dropna().unique().tolist())
    pics = sorted(df['PIC'].dropna().unique().tolist())
    weeks = sorted(df['Week'].dropna().unique().tolist())

    selected_regions = _filter_popover('Region', regions, 'region_filter')
    selected_indices = _filter_popover('KPI INDICES', indices, 'indices_filter')
    selected_pics = _filter_popover('PIC NOS', pics, 'pic_filter')
    selected_weeks = _filter_popover('Week', weeks, 'week_filter')

    selected_regions = selected_regions if selected_regions else regions
    selected_indices = selected_indices if selected_indices else indices
    selected_pics = selected_pics if selected_pics else pics
    selected_weeks = selected_weeks if selected_weeks else weeks

    filtered = df[
        df['Region'].isin(selected_regions)
        & df['KPI INDICES'].isin(selected_indices)
        & df['PIC'].isin(selected_pics)
        & df['Week'].isin(selected_weeks)
    ]
    return filtered


def comparison_by_dimension(df: pd.DataFrame, dim: str, title: str, metric_type: str) -> None:
    grouped = aggregate_metrics(df, [dim], metric_type).sort_values('Increase_abs', ascending=False)
    display_cols = [dim, 'RealKPI', 'NoteUserKPI', 'Increase_abs']
    if metric_type != 'percentage':
        display_cols.append('Increase_pct_vs_real')
    display_df = grouped[display_cols].rename(
        columns={
            'Increase_abs': 'Selisih',
            'Increase_pct_vs_real': 'Selisih (%) vs Real',
        }
    )

    fig = px.bar(
        grouped,
        x=dim,
        y=['RealKPI', 'NoteUserKPI'],
        barmode='group',
        title=title,
        color_discrete_sequence=[REAL_COLOR, NOTE_COLOR],
    )
    fig.update_layout(legend_title_text='Data', plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
    st.plotly_chart(fig, use_container_width=True)

    if metric_type == 'percentage':
        fmt = {
            'RealKPI': '{:.2%}',
            'NoteUserKPI': '{:.2%}',
            'Selisih': '{:.2%}',
        }
    else:
        fmt = {
            'RealKPI': '{:.0f}',
            'NoteUserKPI': '{:.0f}',
            'Selisih': '{:.0f}',
            'Selisih (%) vs Real': '{:.2f}%',
        }

    st.dataframe(display_df.style.format(fmt), use_container_width=True)


def trend_weekly(df: pd.DataFrame, metric_type: str) -> None:
    weekly = aggregate_metrics(df, ['Week'], metric_type)

    fig = px.line(
        weekly,
        x='Week',
        y=['RealKPI', 'NoteUserKPI'],
        markers=True,
        title='Tren Mingguan: RealKPI vs NoteUserKPI',
        color_discrete_sequence=[REAL_COLOR, NOTE_COLOR],
    )
    fig.update_layout(legend_title_text='Data', plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
    st.plotly_chart(fig, use_container_width=True)

    fig_inc = px.bar(
        weekly,
        x='Week',
        y='Increase_abs',
        title='Selisih per Minggu',
        color_discrete_sequence=[INC_COLOR],
    )
    fig_inc.update_layout(showlegend=False, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
    if metric_type == 'percentage':
        fig_inc.update_yaxes(tickformat='.2%')
    st.plotly_chart(fig_inc, use_container_width=True)


def render_analysis_block(group_df: pd.DataFrame, metric_type: str, title_prefix: str) -> None:
    st.markdown('#### Tren Mingguan')
    trend_weekly(group_df, metric_type)

    col_left, col_right = st.columns(2)
    with col_left:
        st.markdown('#### Per KPI INDICES')
        comparison_by_dimension(group_df, 'KPI INDICES', f'{title_prefix}: per KPI INDICES', metric_type)
    with col_right:
        st.markdown('#### Per Region')
        comparison_by_dimension(group_df, 'Region', f'{title_prefix}: per Region', metric_type)

    st.markdown('#### Per PIC NOS')
    comparison_by_dimension(group_df, 'PIC', f'{title_prefix}: per PIC NOS', metric_type)

    st.markdown('#### Detail Data')
    detail_cols = [
        'Region',
        'KPI INDICES',
        'KPI Category',
        'Direction',
        'PIC',
        'Kategori',
        'Week',
        'MetricType',
        'FormatLabel',
        'MetricFormat',
        'RealKPI',
        'NoteUserKPI',
        'Increase_abs',
    ]
    if metric_type != 'percentage':
        detail_cols.append('Increase_pct_vs_real')
    detail_df = group_df[detail_cols].copy()
    detail_df['KPI Category'] = detail_df['KPI Category'].map(category_label)
    detail_df['Direction'] = detail_df['Direction'].map(direction_label)
    if metric_type == 'percentage':
        fmt = {
            'RealKPI': '{:.2%}',
            'NoteUserKPI': '{:.2%}',
            'Increase_abs': '{:.2%}',
        }
    else:
        fmt = {
            'RealKPI': '{:.0f}',
            'NoteUserKPI': '{:.0f}',
            'Increase_abs': '{:.0f}',
            'Increase_pct_vs_real': '{:.2f}%',
        }
    st.dataframe(detail_df.style.format(fmt), use_container_width=True)


def render_category_section(df: pd.DataFrame, category: str) -> None:
    category_display = category_label(category)
    subset = df[df['KPI Category'].eq(category)].copy()

    st.subheader(category_display)

    if subset.empty:
        st.info(f'Tidak ada data {category_display.lower()} pada filter saat ini.')
        return

    groups = (
        subset[['MetricType', 'Direction']]
        .drop_duplicates()
        .sort_values(['MetricType', 'Direction'])
        .to_dict('records')
    )

    group_tabs = st.tabs(
        [
            f"{metric_label(g['MetricType'])} - {direction_label(g['Direction'])}"
            for g in groups
        ]
    )

    for tab, group_info in zip(group_tabs, groups):
        metric_type = group_info['MetricType']
        direction = group_info['Direction']
        metric_display = metric_label(metric_type)
        direction_display = direction_label(direction)
        group_df = subset[
            subset['MetricType'].eq(metric_type)
            & subset['Direction'].eq(direction)
        ].copy()

        with tab:
            title_prefix = f'{category_display} - {metric_display} - {direction_display}'
            st.caption(
                f'{metric_display}. {direction_display}: nilai KPI yang bergerak ke arah ini dianggap lebih baik.'
            )
            render_analysis_block(group_df, metric_type, title_prefix)


def main() -> None:
    inject_styles()

    st.markdown(
        """
        <div class='hero'>
            <h2>Dashboard Perbandingan RealKPI vs NoteUserKPI</h2>
            <p>Pemisahan metrik Percentage dan Absolute agar analisis tidak tercampur</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not DATA_MASTER.exists() or not DATA_FORMATTING.exists():
        st.error('File MASTER DATA.xlsx atau FORMATTING.xlsx tidak ditemukan di folder aplikasi.')
        st.stop()

    df = load_data(DATA_MASTER, DATA_FORMATTING)
    filtered = apply_filters(df)

    if filtered.empty:
        st.warning('Tidak ada data setelah filter diterapkan.')
        st.stop()

    categories = sorted(filtered['KPI Category'].dropna().unique().tolist())
    category_tabs = st.tabs([category_label(category) for category in categories])
    for tab, category in zip(category_tabs, categories):
        with tab:
            render_category_section(filtered, category)


if __name__ == '__main__':
    main()
