import pandas as pd
import plotly.express as px
import streamlit as st
from pathlib import Path
import re

st.set_page_config(page_title='RealKPI vs NoteUserKPI', layout='wide')

DATA_FORMATTING = Path('FORMATTING.xlsx')
GOOGLE_SHEET_ID = '1Mf50M-DCcC0hPXUZd_BMcQNAVKnSudxAUjoNAT_ztys'
GOOGLE_SHEET_GID = '596317790'
DATA_MASTER_URL = f'https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/export?format=csv&gid={GOOGLE_SHEET_GID}'
ID_COLS = ['Region', 'KPI INDICES', 'PIC', 'Kategori']
OVERTIME_OVER_ONTIME_KPIS = {
    'tingkat return',
    'tingkat investigasi kalah',
}

COLOR_BG = '#F2F3F5'
COLOR_CARD = '#FFFFFF'
COLOR_PRIMARY = '#862880'
COLOR_ACCENT = '#EE6825'
COLOR_TEXT = '#1D1D1F'
COLOR_MUTED = '#5F6368'
COLOR_BORDER = '#D7D9DE'

MONTH_ORDER = {
    'jan': 1,
    'january': 1,
    'januari': 1,
    'feb': 2,
    'february': 2,
    'februari': 2,
    'mar': 3,
    'march': 3,
    'maret': 3,
    'apr': 4,
    'april': 4,
    'may': 5,
    'mei': 5,
    'jun': 6,
    'june': 6,
    'juni': 6,
    'jul': 7,
    'july': 7,
    'juli': 7,
    'aug': 8,
    'august': 8,
    'agustus': 8,
    'sep': 9,
    'sept': 9,
    'september': 9,
    'oct': 10,
    'october': 10,
    'oktober': 10,
    'nov': 11,
    'november': 11,
    'dec': 12,
    'december': 12,
    'desember': 12,
}

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
        .hero, .hero *, .hero h2, .hero p {{
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
        }}
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
            -webkit-text-fill-color: #ffffff !important;
        }}
        div[data-baseweb='tag'] div {{
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
        }}
        div[data-baseweb='tag'],
        div[data-baseweb='tag'] *,
        div[data-baseweb='tag'] span,
        div[data-baseweb='tag'] p,
        div[data-baseweb='tag'] svg,
        div[data-baseweb='tag'] button {{
            color: #ffffff !important;
            fill: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
        }}
        div[data-baseweb='tag'] [role='button'],
        div[data-baseweb='tag'] [aria-label],
        div[data-baseweb='tag'] path {{
            color: #ffffff !important;
            fill: #ffffff !important;
            stroke: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
        }}
        .stMultiSelect [data-baseweb='tag'],
        .stMultiSelect [data-baseweb='tag'] *,
        div[data-testid='stMultiSelect'] [data-baseweb='tag'],
        div[data-testid='stMultiSelect'] [data-baseweb='tag'] * {{
            color: #ffffff !important;
            fill: #ffffff !important;
            stroke: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
        }}
        h3, .stSubheader {{
            color: var(--app-text) !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(ttl=3600)
def load_data(master_url: str, formatting_path: Path) -> pd.DataFrame:
    master_df = pd.read_csv(master_url)
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
    overtime_rate_real = over_time.div(on_time.where(on_time.ne(0)))
    overtime_rate_note = current_over_time.div(on_time.where(on_time.ne(0)))

    df['ON TIME'] = on_time
    df['OVER TIME'] = over_time
    df['Current OVER TIME'] = current_over_time
    df['TOTAL OS'] = total_os
    df['Current TOTAL OS'] = current_total_os
    df['RealDenom'] = denom_real
    df['NoteDenom'] = denom_note
    df['CalculationType'] = 'ontime_over_total'
    special_pct_mask = df['KPI INDICES'].astype(str).str.strip().str.lower().isin(OVERTIME_OVER_ONTIME_KPIS)
    df.loc[special_pct_mask, 'CalculationType'] = 'overtime_over_ontime'

    df['RealKPI'] = total_os.astype('float64')
    df['NoteUserKPI'] = current_total_os.astype('float64')
    pct_mask = df['MetricType'].eq('percentage')
    df.loc[pct_mask, 'RealKPI'] = real_pct[pct_mask]
    df.loc[pct_mask, 'NoteUserKPI'] = note_pct[pct_mask]
    special_pct_mask = pct_mask & special_pct_mask
    df.loc[special_pct_mask, 'RealKPI'] = overtime_rate_real[special_pct_mask]
    df.loc[special_pct_mask, 'NoteUserKPI'] = overtime_rate_note[special_pct_mask]

    df['Increase_abs'] = df['NoteUserKPI'] - df['RealKPI']
    df['Increase_pct_vs_real'] = df['Increase_abs'].div(df['RealKPI'].where(df['RealKPI'].ne(0))) * 100
    return df


def aggregate_metrics(df: pd.DataFrame, group_cols: list[str], metric_type: str) -> pd.DataFrame:
    grouped = df.groupby(group_cols, as_index=False) if group_cols else None

    if metric_type == 'percentage':
        is_overtime_rate = (
            'CalculationType' in df.columns
            and df['CalculationType'].eq('overtime_over_ontime').all()
        )
        if grouped is None:
            agg = pd.DataFrame(
                {
                    'ON TIME': [df['ON TIME'].sum()],
                    'OVER TIME': [df['OVER TIME'].sum()],
                    'Current OVER TIME': [df['Current OVER TIME'].sum()],
                    'RealDenom': [df['RealDenom'].sum()],
                    'NoteDenom': [df['NoteDenom'].sum()],
                }
            )
        else:
            agg = grouped[['ON TIME', 'OVER TIME', 'Current OVER TIME', 'RealDenom', 'NoteDenom']].sum()

        if is_overtime_rate:
            agg['RealKPI'] = agg['OVER TIME'].div(agg['ON TIME'].where(agg['ON TIME'].ne(0)))
            agg['NoteUserKPI'] = agg['Current OVER TIME'].div(agg['ON TIME'].where(agg['ON TIME'].ne(0)))
        else:
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


def direction_context(direction: str) -> str:
    normalized = str(direction).strip().lower()
    if normalized == 'higher better':
        return 'Nilai lebih besar menunjukkan performa lebih baik.'
    if normalized == 'lower better':
        return 'Nilai lebih kecil menunjukkan performa lebih baik.'
    return 'Gunakan selisih untuk membaca perubahan antara RealKPI dan NoteUserKPI.'


def add_ranking(df: pd.DataFrame, direction: str) -> pd.DataFrame:
    ranked = df.copy()
    note_ascending = str(direction).strip().lower() == 'lower better'
    ranked['_RankDiffAbs'] = ranked['Increase_abs'].abs()
    ranked = ranked.sort_values(
        ['NoteUserKPI', '_RankDiffAbs'],
        ascending=[note_ascending, True],
        na_position='last',
    ).reset_index(drop=True)
    ranked.insert(0, 'Ranking', range(1, len(ranked) + 1))
    return ranked.drop(columns=['_RankDiffAbs'])


def week_sort_key(week_value: str) -> tuple[int, int, str]:
    text = str(week_value).strip()
    match = re.search(r'week\s*(\d+)\s+([a-zA-Z]+)', text, flags=re.IGNORECASE)
    if not match:
        return (99, 99, text.lower())

    week_number = int(match.group(1))
    month_number = MONTH_ORDER.get(match.group(2).lower(), 99)
    return (month_number, week_number, text.lower())


def sort_weeks(weeks: list[str]) -> list[str]:
    return sorted(weeks, key=week_sort_key)


def apply_metric_axis_format(fig, metric_type: str, value_format: str | None = None):
    if metric_type == 'percentage':
        hover_format = value_format or '.2%'
        fig.update_yaxes(tickformat='.2%')
        fig.update_traces(hovertemplate='%{x}<br>%{fullData.name}: %{y:' + hover_format + '}<extra></extra>')
    return fig


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

    categories = sorted(df['KPI Category'].dropna().unique().tolist())
    regions = sorted(df['Region'].dropna().unique().tolist())
    weeks = sort_weeks(df['Week'].dropna().unique().tolist())

    selected_categories = _filter_popover('Jenis KPI', categories, 'category_filter')
    selected_regions = _filter_popover('Region', regions, 'region_filter')

    active_categories = selected_categories if selected_categories else categories
    active_regions = selected_regions if selected_regions else regions
    pic_source = df[
        df['KPI Category'].isin(active_categories)
        & df['Region'].isin(active_regions)
    ]
    pics = sorted(pic_source['PIC'].dropna().unique().tolist())
    selected_pics = _filter_popover('PIC NOS', pics, 'pic_filter')
    selected_weeks = _filter_popover('Week', weeks, 'week_filter')

    selected_categories = active_categories
    selected_regions = active_regions
    selected_pics = selected_pics if selected_pics else pics
    selected_weeks = selected_weeks if selected_weeks else weeks

    filtered = df[
        df['KPI Category'].isin(selected_categories)
        & df['Region'].isin(selected_regions)
        & df['PIC'].isin(selected_pics)
        & df['Week'].isin(selected_weeks)
    ]
    return filtered


def comparison_by_dimension(
    df: pd.DataFrame,
    dim: str,
    chart_title: str,
    metric_type: str,
    direction: str,
    key_prefix: str,
) -> None:
    grouped = add_ranking(aggregate_metrics(df, [dim], metric_type), direction)
    display_cols = ['Ranking', dim, 'RealKPI', 'NoteUserKPI', 'Increase_abs']
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
        title=chart_title,
        color_discrete_sequence=[REAL_COLOR, NOTE_COLOR],
    )
    fig.update_layout(legend_title_text='Data', plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
    apply_metric_axis_format(fig, metric_type)
    st.plotly_chart(fig, use_container_width=True, key=f'{key_prefix}_{dim}_bar')

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

    st.dataframe(
        display_df.style.format(fmt),
        use_container_width=True,
        hide_index=True,
        key=f'{key_prefix}_{dim}_table',
    )


def trend_weekly(df: pd.DataFrame, metric_type: str, key_prefix: str) -> None:
    weekly = aggregate_metrics(df, ['Week'], metric_type)
    week_order = sort_weeks(weekly['Week'].dropna().unique().tolist())
    weekly['Week'] = pd.Categorical(weekly['Week'], categories=week_order, ordered=True)
    weekly = weekly.sort_values('Week')

    fig = px.line(
        weekly,
        x='Week',
        y=['RealKPI', 'NoteUserKPI'],
        markers=True,
        title='Tren Mingguan: RealKPI vs NoteUserKPI',
        color_discrete_sequence=[REAL_COLOR, NOTE_COLOR],
        category_orders={'Week': week_order},
    )
    fig.update_layout(legend_title_text='Data', plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
    apply_metric_axis_format(fig, metric_type)
    st.plotly_chart(fig, use_container_width=True, key=f'{key_prefix}_weekly_line')

    fig_inc = px.bar(
        weekly,
        x='Week',
        y='Increase_abs',
        title='Selisih per Minggu',
        color_discrete_sequence=[INC_COLOR],
        category_orders={'Week': week_order},
    )
    fig_inc.update_layout(showlegend=False, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
    if metric_type == 'percentage':
        apply_metric_axis_format(fig_inc, metric_type)
    st.plotly_chart(fig_inc, use_container_width=True, key=f'{key_prefix}_weekly_diff')


def kpi_card(title: str, value: str, hint: str = '') -> None:
    st.markdown(
        f"""
        <div class='kpi-card'>
            <div class='kpi-title'>{title}</div>
            <p class='kpi-value'>{value}</p>
            <div class='kpi-hint'>{hint}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def show_single_kpi_cards(df: pd.DataFrame, metric_type: str) -> None:
    summary = aggregate_metrics(df, [], metric_type).iloc[0]
    real_value = summary['RealKPI']
    note_value = summary['NoteUserKPI']
    diff_value = summary['Increase_abs']
    diff_pct = summary['Increase_pct_vs_real']

    if metric_type == 'percentage':
        c1, c2, c3 = st.columns(3)
        with c1:
            kpi_card('RealKPI', f'{real_value:.2%}', 'Nilai baseline')
        with c2:
            kpi_card('NoteUserKPI', f'{note_value:.2%}', 'Nilai setelah note user')
        with c3:
            kpi_card('Selisih', f'{diff_value:.2%}', 'NoteUserKPI - RealKPI')
    else:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            kpi_card('RealKPI', f'{real_value:.0f}', 'Total baseline')
        with c2:
            kpi_card('NoteUserKPI', f'{note_value:.0f}', 'Total setelah note user')
        with c3:
            kpi_card('Selisih', f'{diff_value:.0f}', 'NoteUserKPI - RealKPI')
        with c4:
            kpi_card('Selisih vs Real', f'{diff_pct:.2f}%', 'Proporsi selisih terhadap RealKPI')


def render_analysis_block(group_df: pd.DataFrame, metric_type: str, direction: str, key_prefix: str) -> None:
    show_single_kpi_cards(group_df, metric_type)
    direction_note = direction_context(direction)

    st.markdown('#### Tren Mingguan')
    st.caption(f'Membandingkan RealKPI dan NoteUserKPI dari minggu ke minggu. {direction_note}')
    trend_weekly(group_df, metric_type, key_prefix)

    col_left, col_right = st.columns(2)
    with col_left:
        st.markdown('#### Per KPI INDICES')
        st.caption(f'Ringkasan nilai KPI pada indeks yang sedang dibuka. {direction_note}')
        comparison_by_dimension(
            group_df,
            'KPI INDICES',
            'Ringkasan RealKPI vs NoteUserKPI',
            metric_type,
            direction,
            key_prefix,
        )
    with col_right:
        st.markdown('#### Per Region')
        st.caption(f'Membandingkan nilai KPI antar region untuk melihat area yang paling terdampak note user. {direction_note}')
        comparison_by_dimension(
            group_df,
            'Region',
            'Perbandingan antar Region',
            metric_type,
            direction,
            key_prefix,
        )

    st.markdown('#### Per PIC NOS')
    st.caption(f'Membandingkan nilai KPI antar PIC NOS untuk melihat perubahan terbesar setelah note user. {direction_note}')
    comparison_by_dimension(group_df, 'PIC', 'Perbandingan antar PIC NOS', metric_type, direction, key_prefix)

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
        'CalculationType',
        'FormatLabel',
        'MetricFormat',
        'RealKPI',
        'NoteUserKPI',
        'Increase_abs',
    ]
    if metric_type != 'percentage':
        detail_cols.append('Increase_pct_vs_real')
    detail_df = group_df[detail_cols].copy()
    detail_week_order = sort_weeks(detail_df['Week'].dropna().unique().tolist())
    detail_df['Week'] = pd.Categorical(detail_df['Week'], categories=detail_week_order, ordered=True)
    detail_df = detail_df.sort_values(['Week', 'Region', 'PIC']).reset_index(drop=True)
    detail_df = add_ranking(detail_df, direction)
    detail_df['Week'] = detail_df['Week'].astype(str)
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
    st.dataframe(
        detail_df.style.format(fmt),
        use_container_width=True,
        hide_index=True,
        key=f'{key_prefix}_detail_table',
    )


def main() -> None:
    inject_styles()

    st.markdown(
        """
        <div class='hero'>
            <h2>Dashboard Perbandingan RealKPI vs NoteUserKPI</h2>
            <p>Analisis tren dan selisih KPI NOS per indeks, region, PIC, dan minggu</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not DATA_FORMATTING.exists():
        st.error('File FORMATTING.xlsx tidak ditemukan di folder aplikasi.')
        st.stop()

    try:
        df = load_data(DATA_MASTER_URL, DATA_FORMATTING)
    except Exception as exc:
        st.error(f'Gagal membaca data dari Google Sheets: {exc}')
        st.stop()
    filtered = apply_filters(df)

    if filtered.empty:
        st.warning('Tidak ada data setelah filter diterapkan.')
        st.stop()

    indices = sorted(filtered['KPI INDICES'].dropna().unique().tolist())
    kpi_tabs = st.tabs(indices)

    for i, (tab, kpi_index) in enumerate(zip(kpi_tabs, indices)):
        with tab:
            kpi_df = filtered[filtered['KPI INDICES'].eq(kpi_index)].copy()
            meta = kpi_df.iloc[0]
            metric_type = meta['MetricType']

            st.subheader(kpi_index)
            st.caption(
                f"{category_label(meta['KPI Category'])}. "
                f"{metric_label(metric_type)}. "
                f"{direction_label(meta['Direction'])}. "
                f"{direction_context(meta['Direction'])}"
            )
            render_analysis_block(kpi_df, metric_type, meta['Direction'], key_prefix=f'kpi_{i}')


if __name__ == '__main__':
    main()
