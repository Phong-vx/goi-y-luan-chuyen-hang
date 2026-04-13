import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
from pathlib import Path
import pandas as pd
import numpy as np
from datetime import datetime
import math
from PIL import Image, ImageTk
import io

# ─── Constants ────────────────────────────────────────────────────────────────
APP_TITLE = "Phân Tích Luân Chuyển Hàng Hoá"
PRIMARY   = "#1565C0"
PRIMARY_L = "#1976D2"
ACCENT    = "#0D47A1"
BG        = "#F5F7FA"
CARD_BG   = "#FFFFFF"
TEXT_DARK = "#1A1A2E"
TEXT_MID  = "#4A4A68"
TEXT_LIGHT= "#9E9E9E"
SUCCESS   = "#2E7D32"
WARNING   = "#F57F17"
DANGER    = "#C62828"
BORDER    = "#E0E6EF"


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


def get_region(store_name: str) -> str:
    """Detect HCM / HN / OTHER from store name."""
    name = str(store_name).upper()
    if 'HCM' in name:
        return 'HCM'
    if 'HN' in name:
        return 'HN'
    return 'OTHER'


# ─── Logic ────────────────────────────────────────────────────────────────────

def calculate(df_sales: pd.DataFrame, df_inv: pd.DataFrame,
              target_months: float, slow_pct: float,
              selected_stores=None,
              allow_hcm_hn: bool = True,
              min_fill_avg: float = 0.4):
    """
    Returns (df_fill, df_transfer, warnings)
    selected_stores: list of store names to include; None = all
    allow_hcm_hn:   if False, block cross-region (HCM<->HN) transfers
    min_fill_avg:   AvgMonthly <= this value => skip fill
    """
    warnings = []

    df_sales = df_sales.copy()
    df_inv   = df_inv.copy()

    df_sales.columns = df_sales.columns.str.strip()
    df_inv.columns   = df_inv.columns.str.strip()

    df_sales['Date']     = pd.to_datetime(df_sales['Date'], errors='coerce')
    df_sales['Quantity'] = pd.to_numeric(df_sales['Quantity'], errors='coerce').fillna(0)
    df_inv['Số lượng']   = pd.to_numeric(df_inv['Số lượng'],   errors='coerce').fillna(0)

    # ── filter retail sales ──────────────────────────────────────────────────
    retail_sales = df_sales[df_sales['Sale Team'].str.strip() == 'Bán Lẻ'].copy()
    if retail_sales.empty:
        warnings.append("Không tìm thấy dữ liệu bán lẻ (Sale Team = 'Bán Lẻ').")

    if selected_stores:
        retail_sales = retail_sales[retail_sales['location Name'].isin(selected_stores)]

    # ── last 3 months window ─────────────────────────────────────────────────
    if not retail_sales.empty:
        max_date  = retail_sales['Date'].max()
        min_3m    = max_date - pd.DateOffset(months=3)
        sales_3m  = retail_sales[retail_sales['Date'] >= min_3m].copy()
        actual_months = min(3.0, max(
            (max_date - retail_sales['Date'].min()).days / 30.44, 0.1
        ))
    else:
        sales_3m      = retail_sales
        actual_months = 3.0

    # ── velocity per store × SKU ─────────────────────────────────────────────
    vel = (
        sales_3m
        .groupby(['location Name', 'SKU'], as_index=False)['Quantity']
        .sum()
        .rename(columns={'location Name': 'Store', 'Quantity': 'Sold3M'})
    )
    vel['AvgMonthly'] = vel['Sold3M'] / actual_months

    # ── Revenue column lookup ─────────────────────────────────────────────────
    _REV_COLS = ['Revenue', 'Amount', 'Doanh thu', 'Doanh Thu', 'Thành tiền',
                 'Net Amount', 'Total Amount', 'Giá trị', 'Total', 'Subtotal']
    _rev_col  = next((c for c in _REV_COLS if c in retail_sales.columns), None)
    if _rev_col:
        retail_sales[_rev_col] = pd.to_numeric(retail_sales[_rev_col], errors='coerce').fillna(0)
        sales_3m[_rev_col]     = pd.to_numeric(sales_3m[_rev_col],     errors='coerce').fillna(0)
        all_time = (
            retail_sales
            .groupby(['location Name', 'SKU'], as_index=False)
            .agg({'Quantity': 'sum', _rev_col: 'sum'})
            .rename(columns={'location Name': 'Store', 'Quantity': 'SoldAll', _rev_col: 'RevAll'})
        )
        rev3m_df = (
            sales_3m
            .groupby(['location Name', 'SKU'], as_index=False)[_rev_col]
            .sum()
            .rename(columns={'location Name': 'Store', _rev_col: 'Rev3M'})
        )
    else:
        all_time = (
            retail_sales
            .groupby(['location Name', 'SKU'], as_index=False)['Quantity']
            .sum()
            .rename(columns={'location Name': 'Store', 'Quantity': 'SoldAll'})
        )
        all_time['RevAll'] = 0
        rev3m_df = vel[['Store', 'SKU']].copy()
        rev3m_df['Rev3M'] = 0
    all_time['SKU']  = all_time['SKU'].astype(str)
    rev3m_df['SKU']  = rev3m_df['SKU'].astype(str)

    prod_names = (
        retail_sales[['SKU', 'Product Item']]
        .drop_duplicates('SKU')
        .rename(columns={'Product Item': 'ProductName'})
    )

    # ── SKU → Product attributes mapping ─────────────────────────────────────
    _PROD_ATTR_COLS = ['BRAND', 'Category', 'Model', 'Sub Category', 'Color', 'Frame Size']
    _attr_lookup = {
        'BRAND':        ['BRAND', 'Brand', 'Thương hiệu'],
        'Category':     ['Category', 'Danh mục', 'Danh Mục', 'Product Category',
                         'Nhóm hàng', 'Nhóm SP', 'Nhóm'],
        'Model':        ['Model', 'Mô hình', 'Product Model'],
        'Sub Category': ['Sub Category', 'SubCategory', 'Danh mục phụ', 'Nhóm con'],
        'Color':        ['Color', 'Colour', 'Màu sắc', 'Màu'],
        'Frame Size':   ['Frame Size', 'FrameSize', 'Size', 'Kích thước khung'],
    }
    _col_map = {}   # target_attr → actual column name in df_sales
    for _attr, _candidates in _attr_lookup.items():
        _found = next((c for c in _candidates if c in df_sales.columns), None)
        if _found:
            _col_map[_attr] = _found

    if _col_map:
        _src = ['SKU'] + list(dict.fromkeys(_col_map.values()))   # deduplicate
        prod_attrs = df_sales[_src].drop_duplicates('SKU').copy()
        prod_attrs = prod_attrs.rename(columns={v: k for k, v in _col_map.items()})
    else:
        prod_attrs = pd.DataFrame({'SKU': df_sales['SKU'].unique()})
    prod_attrs['SKU'] = prod_attrs['SKU'].astype(str)
    for _attr in _PROD_ATTR_COLS:
        if _attr not in prod_attrs.columns:
            prod_attrs[_attr] = ''

    # ── retail inventory ─────────────────────────────────────────────────────
    inv_retail = df_inv[df_inv['Địa điểm/Team'].str.strip() == 'Bán Lẻ'].copy()
    inv_retail = inv_retail.rename(columns={
        'Địa điểm/Tên hiển thị': 'Store',
        'Sản phẩm/Mã nội bộ':   'SKU',
        'Sản phẩm/Tên hiển thị': 'ProductName',
        'Số lượng':              'StoreQty',
    })
    inv_retail['SKU']      = inv_retail['SKU'].astype(str)
    inv_retail['StoreQty'] = inv_retail['StoreQty'].fillna(0)

    if selected_stores:
        inv_retail = inv_retail[inv_retail['Store'].isin(selected_stores)]

    store_inv = (
        inv_retail
        .groupby(['Store', 'SKU'], as_index=False)['StoreQty']
        .sum()
    )

    # ── warehouse inventory ──────────────────────────────────────────────────
    inv_kho = df_inv[df_inv['Địa điểm/Team'].str.strip() == 'Kho'].copy()
    inv_kho = inv_kho.rename(columns={
        'Sản phẩm/Mã nội bộ': 'SKU',
        'Số lượng':            'KhoQty',
    })
    inv_kho['SKU']    = inv_kho['SKU'].astype(str)
    inv_kho['KhoQty'] = inv_kho['KhoQty'].fillna(0)

    kho_total = inv_kho.groupby('SKU', as_index=False)['KhoQty'].sum()

    # ── SHEET 1: fill from warehouse ─────────────────────────────────────────
    fill = vel.copy()
    fill['SKU'] = fill['SKU'].astype(str)

    # Drop rows below min_fill_avg threshold
    fill = fill[fill['AvgMonthly'] > min_fill_avg].copy()

    fill = fill.merge(store_inv, on=['Store', 'SKU'], how='left')
    fill['StoreQty'] = fill['StoreQty'].fillna(0)
    fill = fill.merge(kho_total, on='SKU', how='left')
    fill['KhoQty'] = fill['KhoQty'].fillna(0)
    prod_names['SKU'] = prod_names['SKU'].astype(str)
    fill = fill.merge(prod_names, on='SKU', how='left')

    fill['TargetStock']   = np.ceil(fill['AvgMonthly'] * target_months).astype(int)
    fill['SuggestedFill'] = (fill['TargetStock'] - fill['StoreQty']).clip(lower=0)
    fill['WeeksOfStock']  = np.where(
        fill['AvgMonthly'] > 0,
        (fill['StoreQty'] / fill['AvgMonthly'] * 4.33).round(1),
        999.0
    )

    fill = fill.sort_values('AvgMonthly', ascending=False)
    remaining_kho = kho_total.set_index('SKU')['KhoQty'].to_dict()
    final_fills = []
    for _, row in fill.iterrows():
        sku   = row['SKU']
        avail = remaining_kho.get(sku, 0)
        qty   = min(int(row['SuggestedFill']), int(avail))
        final_fills.append(qty)
        remaining_kho[sku] = max(0, avail - qty)
    fill['FinalFill'] = final_fills

    # Merge all-time + 3-month revenue into fill
    fill = fill.merge(all_time[['Store', 'SKU', 'SoldAll', 'RevAll']], on=['Store', 'SKU'], how='left')
    fill = fill.merge(rev3m_df[['Store', 'SKU', 'Rev3M']],             on=['Store', 'SKU'], how='left')
    fill['SoldAll'] = fill['SoldAll'].fillna(0).astype(int)
    fill['RevAll']  = fill['RevAll'].fillna(0)
    fill['Rev3M']   = fill['Rev3M'].fillna(0)

    df_fill = fill[fill['FinalFill'] > 0].copy()
    df_fill = df_fill[[
        'Store', 'SKU', 'ProductName',
        'SoldAll', 'RevAll', 'Sold3M', 'Rev3M',
        'AvgMonthly', 'WeeksOfStock',
        'StoreQty', 'TargetStock', 'SuggestedFill',
        'KhoQty', 'FinalFill',
    ]].rename(columns={
        'Store':         'Cửa hàng',
        'SKU':           'Mã hàng',
        'ProductName':   'Tên sản phẩm',
        'SoldAll':       'Tổng bán toàn TG',
        'RevAll':        'Doanh thu toàn TG',
        'Sold3M':        'Đã bán 3T',
        'Rev3M':         'Doanh thu 3T',
        'AvgMonthly':    'Sức bán TB/tháng',
        'WeeksOfStock':  'Tuần tồn kho',
        'StoreQty':      'Tồn cửa hàng',
        'TargetStock':   'Mức tồn mục tiêu',
        'SuggestedFill': 'Đề xuất fill',
        'KhoQty':        'Tồn kho',
        'FinalFill':     'Fill thực tế',
    })
    df_fill['Sức bán TB/tháng'] = df_fill['Sức bán TB/tháng'].round(2)
    df_fill = df_fill.sort_values(['Cửa hàng', 'Sức bán TB/tháng'], ascending=[True, False])

    # Product attributes + Trạng thái cho Sheet 1
    df_fill = (df_fill
               .merge(prod_attrs[['SKU'] + _PROD_ATTR_COLS], left_on='Mã hàng', right_on='SKU', how='left')
               .drop(columns=['SKU']))
    for _attr in _PROD_ATTR_COLS:
        df_fill[_attr] = df_fill[_attr].fillna('')
    df_fill['Trạng thái'] = df_fill['Fill thực tế'].apply(
        lambda x: 'Cần Fill hàng' if x >= 1 else 'Đủ hàng'
    )
    _fill_cols = (['Cửa hàng', 'Mã hàng', 'Tên sản phẩm'] + _PROD_ATTR_COLS +
                  ['Tổng bán toàn TG', 'Doanh thu toàn TG', 'Đã bán 3T', 'Doanh thu 3T',
                   'Sức bán TB/tháng', 'Tuần tồn kho',
                   'Tồn cửa hàng', 'Mức tồn mục tiêu', 'Đề xuất fill',
                   'Tồn kho', 'Fill thực tế', 'Trạng thái'])
    df_fill = df_fill[[c for c in _fill_cols if c in df_fill.columns]]

    # ── SHEET 2: transfers between stores ────────────────────────────────────
    relevant_skus = set(vel['SKU'].astype(str).unique()) | set(
        store_inv[store_inv['StoreQty'] > 0]['SKU'].astype(str).unique()
    )

    grid_rows = []
    for sku in relevant_skus:
        v_sku = vel[vel['SKU'].astype(str) == sku].set_index('Store')['AvgMonthly'].to_dict()
        i_sku = store_inv[store_inv['SKU'].astype(str) == sku].set_index('Store')['StoreQty'].to_dict()
        all_s = set(v_sku.keys()) | set(i_sku.keys())
        for store in all_s:
            grid_rows.append({
                'SKU':        sku,
                'Store':      store,
                'AvgMonthly': v_sku.get(store, 0.0),
                'StoreQty':   i_sku.get(store, 0),
            })

    _empty_transfer = pd.DataFrame(columns=(
        ['Mã hàng', 'Tên sản phẩm'] + _PROD_ATTR_COLS +
        ['Cửa hàng gửi', 'Tồn gửi', 'Sức bán gửi TB/tháng',
         'Cửa hàng nhận', 'Tồn nhận', 'Sức bán nhận TB/tháng',
         'Đề xuất luân chuyển', 'Trạng thái']
    ))

    if not grid_rows:
        # build minimal summary from fill only
        if not df_fill.empty:
            _summ = (df_fill
                     .groupby(['Cửa hàng', 'Category'], as_index=False)['Fill thực tế']
                     .sum()
                     .rename(columns={'Fill thực tế': 'Fill từ kho'}))
            _summ['Chuyển đi'] = 0
            _summ['Nhận đến']  = 0
            _summ = _summ.sort_values(['Cửa hàng', 'Category'])
        else:
            _summ = pd.DataFrame(columns=['Cửa hàng', 'Category',
                                          'Fill từ kho', 'Chuyển đi', 'Nhận đến'])
        return df_fill, _empty_transfer, _summ, warnings

    grid = pd.DataFrame(grid_rows)
    sku_avg_vel = grid.groupby('SKU')['AvgMonthly'].mean().to_dict()

    transfer_rows = []
    for sku, grp in grid.groupby('SKU'):
        net_avg   = sku_avg_vel.get(sku, 0)
        threshold = net_avg * (slow_pct / 100.0)

        slow_stores = grp[(grp['StoreQty'] > 0) & (grp['AvgMonthly'] <= threshold)].copy()
        fast_stores = grp[(grp['AvgMonthly'] > threshold)].copy()

        fast_stores = fast_stores.copy()
        fast_stores['Need'] = (
            np.ceil(fast_stores['AvgMonthly'] * target_months).astype(int)
            - fast_stores['StoreQty']
        )
        fast_stores = fast_stores[fast_stores['Need'] > 0].sort_values('Need', ascending=False)

        if slow_stores.empty or fast_stores.empty:
            continue

        slow_avail = slow_stores.set_index('Store')['StoreQty'].to_dict()
        slow_vel   = slow_stores.set_index('Store')['AvgMonthly'].to_dict()

        for _, fast_row in fast_stores.iterrows():
            need = int(fast_row['Need'])
            if need <= 0:
                continue
            for slow_store in list(slow_avail.keys()):
                avail = slow_avail[slow_store]
                if avail <= 0:
                    continue
                # ── Region rule ──────────────────────────────────────────────
                if not allow_hcm_hn:
                    fr = get_region(fast_row['Store'])
                    sr = get_region(slow_store)
                    if fr != sr and fr != 'OTHER' and sr != 'OTHER':
                        continue  # block cross-region
                qty = min(need, avail)
                transfer_rows.append({
                    'SKU':            sku,
                    'FastStore':      fast_row['Store'],
                    'FastAvgMonthly': round(fast_row['AvgMonthly'], 2),
                    'FastStoreQty':   int(fast_row['StoreQty']),
                    'SlowStore':      slow_store,
                    'SlowAvgMonthly': round(slow_vel[slow_store], 2),
                    'SlowStoreQty':   int(avail),
                    'TransferQty':    qty,
                })
                slow_avail[slow_store] -= qty
                need -= qty
                if need <= 0:
                    break

    if transfer_rows:
        df_transfer = pd.DataFrame(transfer_rows)
        df_transfer['SKU'] = df_transfer['SKU'].astype(str)
        pn = prod_names.copy()
        pn['SKU'] = pn['SKU'].astype(str)
        df_transfer = df_transfer.merge(pn, on='SKU', how='left')
        df_transfer = df_transfer[[
            'SKU', 'ProductName',
            'SlowStore', 'SlowStoreQty', 'SlowAvgMonthly',
            'FastStore',  'FastStoreQty',  'FastAvgMonthly',
            'TransferQty',
        ]].rename(columns={
            'SKU':            'Mã hàng',
            'ProductName':    'Tên sản phẩm',
            'SlowStore':      'Cửa hàng gửi',
            'SlowStoreQty':   'Tồn gửi',
            'SlowAvgMonthly': 'Sức bán gửi TB/tháng',
            'FastStore':      'Cửa hàng nhận',
            'FastStoreQty':   'Tồn nhận',
            'FastAvgMonthly': 'Sức bán nhận TB/tháng',
            'TransferQty':    'Đề xuất luân chuyển',
        })
        df_transfer = df_transfer.sort_values(['Mã hàng', 'Cửa hàng gửi'])

        # Product attributes + Trạng thái cho Sheet 2
        df_transfer = (df_transfer
                       .merge(prod_attrs[['SKU'] + _PROD_ATTR_COLS], left_on='Mã hàng', right_on='SKU', how='left')
                       .drop(columns=['SKU']))
        for _attr in _PROD_ATTR_COLS:
            df_transfer[_attr] = df_transfer[_attr].fillna('')
        df_transfer['Trạng thái'] = df_transfer['Đề xuất luân chuyển'].apply(
            lambda x: 'Cần Luân Chuyển' if x >= 1 else 'Đủ hàng'
        )
        _xfer_cols = (['Mã hàng', 'Tên sản phẩm'] + _PROD_ATTR_COLS +
                      ['Cửa hàng gửi', 'Tồn gửi', 'Sức bán gửi TB/tháng',
                       'Cửa hàng nhận', 'Tồn nhận', 'Sức bán nhận TB/tháng',
                       'Đề xuất luân chuyển', 'Trạng thái'])
        df_transfer = df_transfer[[c for c in _xfer_cols if c in df_transfer.columns]]
    else:
        df_transfer = _empty_transfer

    # ── SHEET 3: summary by store × category ─────────────────────────────────
    fill_agg = (df_fill
                .groupby(['Cửa hàng', 'Category'], as_index=False)['Fill thực tế']
                .sum()
                .rename(columns={'Fill thực tế': 'Fill từ kho'}))

    if not df_transfer.empty:
        sent_agg = (df_transfer
                    .groupby(['Cửa hàng gửi', 'Category'], as_index=False)['Đề xuất luân chuyển']
                    .sum()
                    .rename(columns={'Cửa hàng gửi': 'Cửa hàng',
                                     'Đề xuất luân chuyển': 'Chuyển đi'}))
        recv_agg = (df_transfer
                    .groupby(['Cửa hàng nhận', 'Category'], as_index=False)['Đề xuất luân chuyển']
                    .sum()
                    .rename(columns={'Cửa hàng nhận': 'Cửa hàng',
                                     'Đề xuất luân chuyển': 'Nhận đến'}))
    else:
        sent_agg = pd.DataFrame(columns=['Cửa hàng', 'Category', 'Chuyển đi'])
        recv_agg = pd.DataFrame(columns=['Cửa hàng', 'Category', 'Nhận đến'])

    _pairs = set()
    for _d in (fill_agg, sent_agg, recv_agg):
        if not _d.empty:
            _pairs.update(zip(_d['Cửa hàng'], _d['Category']))

    if _pairs:
        _base = pd.DataFrame(list(_pairs), columns=['Cửa hàng', 'Category'])
        df_summary = (_base
                      .merge(fill_agg, on=['Cửa hàng', 'Category'], how='left')
                      .merge(sent_agg, on=['Cửa hàng', 'Category'], how='left')
                      .merge(recv_agg, on=['Cửa hàng', 'Category'], how='left'))
        for _c in ['Fill từ kho', 'Chuyển đi', 'Nhận đến']:
            df_summary[_c] = df_summary[_c].fillna(0).astype(int)
        df_summary = df_summary.sort_values(['Cửa hàng', 'Category'])
    else:
        df_summary = pd.DataFrame(columns=['Cửa hàng', 'Category',
                                           'Fill từ kho', 'Chuyển đi', 'Nhận đến'])

    return df_fill, df_transfer, df_summary, warnings


def export_excel(df_fill: pd.DataFrame, df_transfer: pd.DataFrame,
                 df_summary: pd.DataFrame, save_path: str):
    with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
        workbook = writer.book

        def _f(props):
            base = {'border': 1, 'valign': 'vcenter'}
            return workbook.add_format({**base, **props})

        title_fmt = workbook.add_format({
            'bold': True, 'font_size': 13,
            'bg_color': '#0D47A1', 'font_color': '#FFFFFF',
            'align': 'center', 'valign': 'vcenter',
        })
        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#1565C0', 'font_color': '#FFFFFF',
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        })
        c_odd, c_even = '#FFFFFF', '#DCEEFB'
        cell = [_f({'bg_color': c_odd,  'align': 'left'}),
                _f({'bg_color': c_even, 'align': 'left'})]
        num  = [_f({'bg_color': c_odd,  'align': 'center', 'num_format': '#,##0'}),
                _f({'bg_color': c_even, 'align': 'center', 'num_format': '#,##0'})]
        dec  = [_f({'bg_color': c_odd,  'align': 'center', 'num_format': '#,##0.00'}),
                _f({'bg_color': c_even, 'align': 'center', 'num_format': '#,##0.00'})]
        hl_fill = [_f({'bold': True, 'bg_color': '#90CAF9', 'align': 'center', 'num_format': '#,##0'}),
                   _f({'bold': True, 'bg_color': '#BBDEFB', 'align': 'center', 'num_format': '#,##0'})]
        hl_xfer = [_f({'bold': True, 'bg_color': '#FFD54F', 'align': 'center', 'num_format': '#,##0'}),
                   _f({'bold': True, 'bg_color': '#FFF8E1', 'align': 'center', 'num_format': '#,##0'})]
        # Trạng thái: Cần Fill/Cần Luân Chuyển = cam đỏ, Đủ hàng = xanh lá
        status_need = [
            workbook.add_format({'bold': True, 'bg_color': '#FFCCBC', 'font_color': '#BF360C',
                                  'border': 1, 'align': 'center', 'valign': 'vcenter'}),
            workbook.add_format({'bold': True, 'bg_color': '#FFE0B2', 'font_color': '#BF360C',
                                  'border': 1, 'align': 'center', 'valign': 'vcenter'}),
        ]
        status_ok = [
            workbook.add_format({'bold': True, 'bg_color': '#C8E6C9', 'font_color': '#1B5E20',
                                  'border': 1, 'align': 'center', 'valign': 'vcenter'}),
            workbook.add_format({'bold': True, 'bg_color': '#DCEDC8', 'font_color': '#1B5E20',
                                  'border': 1, 'align': 'center', 'valign': 'vcenter'}),
        ]
        sub_text  = workbook.add_format({
            'bold': True, 'italic': True,
            'bg_color': '#1E88E5', 'font_color': '#FFFFFF',
            'border': 1, 'align': 'left', 'valign': 'vcenter',
        })
        sub_num   = workbook.add_format({
            'bold': True, 'italic': True,
            'bg_color': '#1E88E5', 'font_color': '#FFFFFF',
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'num_format': '#,##0',
        })
        sub_blank = workbook.add_format({'bg_color': '#1E88E5', 'border': 1})

        def write_sheet(df, sheet_name, title_text, group_col, sum_cols,
                        highlight_col=None, hl_fmts=None, show_subtotal=True):
            n_cols   = len(df.columns)
            col_list = list(df.columns)
            ws       = workbook.add_worksheet(sheet_name)
            ws.merge_range(0, 0, 0, n_cols - 1, title_text, title_fmt)
            ws.set_row(0, 30)
            for ci, name in enumerate(col_list):
                ws.write(1, ci, name, header_fmt)
            ws.set_row(1, 38)
            ws.freeze_panes(2, 3)

            ri = 2
            for group_val, grp in df.groupby(group_col, sort=False):
                sums = {c: 0 for c in sum_cols}
                alt  = 0
                for _, row in grp.iterrows():
                    p = alt % 2
                    for ci, col in enumerate(col_list):
                        val = row[col]
                        if val is None or (isinstance(val, float) and np.isnan(val)):
                            val = ''
                        if col in sum_cols and isinstance(val, (int, float, np.integer, np.floating)):
                            sums[col] += val
                        if col == 'Trạng thái':
                            is_need = str(val).startswith('Cần')
                            fmt = status_need[p] if is_need else status_ok[p]
                        elif col == highlight_col and hl_fmts:
                            fmt = hl_fmts[p]
                        elif isinstance(val, float):
                            fmt = dec[p]
                        elif isinstance(val, (int, np.integer)):
                            fmt = num[p]
                        else:
                            fmt = cell[p]
                        ws.write(ri, ci, val if val != '' else '', fmt)
                    ws.set_row(ri, 18)
                    ri  += 1
                    alt += 1
                if show_subtotal:
                    for ci, col in enumerate(col_list):
                        if ci == 0:
                            ws.write(ri, ci, f'Tổng  {group_val}', sub_text)
                        elif col in sum_cols:
                            ws.write(ri, ci, sums[col], sub_num)
                        else:
                            ws.write(ri, ci, '', sub_blank)
                    ws.set_row(ri, 20)
                    ri += 1

            for ci, col in enumerate(col_list):
                if df.empty:
                    max_len = len(str(col))
                else:
                    max_len = max(len(str(col)), df[col].astype(str).str.len().max())
                ws.set_column(ci, ci, min(max_len + 2, 48))

        write_sheet(
            df_fill, 'Fill từ kho',
            title_text    = 'PHÂN TÍCH FILL HÀNG TỪ KHO',
            group_col     = 'Cửa hàng',
            sum_cols      = ['Tổng bán toàn TG', 'Doanh thu toàn TG',
                             'Đã bán 3T', 'Doanh thu 3T',
                             'Tồn cửa hàng', 'Mức tồn mục tiêu',
                             'Đề xuất fill', 'Fill thực tế'],
            highlight_col = 'Fill thực tế',
            hl_fmts       = hl_fill,
            show_subtotal = False,
        )
        write_sheet(
            df_transfer, 'Luân chuyển cửa hàng',
            title_text    = 'PHÂN TÍCH LUÂN CHUYỂN HÀNG HOÁ GIỮA CÁC CỬA HÀNG',
            group_col     = 'Cửa hàng gửi',
            sum_cols      = ['Tồn gửi', 'Đề xuất luân chuyển'],
            highlight_col = 'Đề xuất luân chuyển',
            hl_fmts       = hl_xfer,
        )
        write_sheet(
            df_summary, 'Tổng hợp',
            title_text    = 'TỔNG HỢP PHÂN TÍCH LUÂN CHUYỂN HÀNG HOÁ',
            group_col     = 'Cửa hàng',
            sum_cols      = ['Fill từ kho', 'Chuyển đi', 'Nhận đến'],
            highlight_col = None,
        )


# ─── CheckListbox ─────────────────────────────────────────────────────────────

class CheckListbox(tk.Frame):
    """Scrollable list of Checkbutton widgets with select/deselect-all and live filter."""

    def __init__(self, parent, height=200, **kwargs):
        super().__init__(parent, bg='#FAFAFA',
                         highlightbackground=BORDER, highlightthickness=1,
                         **kwargs)
        self.configure(height=height)
        self.pack_propagate(False)

        self._canvas = tk.Canvas(self, bg='#FAFAFA', highlightthickness=0)
        _vsb = ttk.Scrollbar(self, orient='vertical', command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=_vsb.set)
        _vsb.pack(side='right', fill='y')
        self._canvas.pack(side='left', fill='both', expand=True)

        self._inner  = tk.Frame(self._canvas, bg='#FAFAFA')
        self._win_id = self._canvas.create_window((0, 0), window=self._inner, anchor='nw')

        self._inner.bind('<Configure>',
            lambda e: self._canvas.configure(scrollregion=self._canvas.bbox('all')))
        self._canvas.bind('<Configure>',
            lambda e: self._canvas.itemconfig(self._win_id, width=e.width))

        # Shared scroll handler — also reused in _render for each Checkbutton
        def _scroll(e):
            delta = e.delta if e.delta else (-120 if getattr(e, 'num', 0) == 4 else 120)
            self._canvas.yview_scroll(int(-1 * delta / 120), 'units')
        self._scroll_cmd = _scroll
        self._canvas.bind('<MouseWheel>',  self._scroll_cmd)
        self._canvas.bind('<Button-4>',    self._scroll_cmd)   # Linux scroll up
        self._canvas.bind('<Button-5>',    self._scroll_cmd)   # Linux scroll down
        self._inner.bind('<MouseWheel>',   self._scroll_cmd)

        self._var_map:   dict = {}   # {item_name: BooleanVar} — full state
        self._all_items: list = []

    def set_items(self, items: list, select_all: bool = True):
        self._all_items = list(items)
        self._var_map   = {item: tk.BooleanVar(value=select_all) for item in items}
        self._render(self._all_items)

    def filter(self, text: str):
        q = text.strip().lower()
        visible = self._all_items if not q else [
            i for i in self._all_items if q in i.lower()
        ]
        self._render(visible)

    def _render(self, items: list):
        for w in self._inner.winfo_children():
            w.destroy()
        for item in items:
            var = self._var_map[item]
            cb = tk.Checkbutton(
                self._inner, text=item, variable=var,
                bg='#FAFAFA', fg=TEXT_DARK,
                font=('Segoe UI', 9),
                activebackground='#E3F2FD',
                selectcolor='white', anchor='w',
                relief='flat', bd=0,
            )
            cb.pack(fill='x', padx=6, pady=1)
            cb.bind('<MouseWheel>', self._scroll_cmd)
            cb.bind('<Button-4>',   self._scroll_cmd)
            cb.bind('<Button-5>',   self._scroll_cmd)
        self._canvas.yview_moveto(0)

    def select_all(self):
        for var in self._var_map.values():
            var.set(True)

    def deselect_all(self):
        for var in self._var_map.values():
            var.set(False)

    def get_selected(self) -> list:
        return [item for item, var in self._var_map.items() if var.get()]


# ─── GUI ──────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.resizable(True, True)
        self.minsize(900, 750)
        self.configure(bg=BG)

        # state
        self.sales_path    = tk.StringVar()
        self.inv_path      = tk.StringVar()
        self.target_months = tk.DoubleVar(value=1.5)
        self.slow_pct      = tk.DoubleVar(value=50.0)
        self.min_fill_avg  = tk.DoubleVar(value=0.4)
        self.allow_hcm_hn  = tk.BooleanVar(value=False)
        self.status_text   = tk.StringVar(value="Chọn file và bấm Đọc File")

        self._physical_stores: list = []
        self._shopee_stores:   list = []
        self.df_sales_raw = None
        self.df_inv_raw   = None

        self._build_ui()
        self._center()

        # Enable "Đọc File" button when both files are selected
        self.sales_path.trace_add('write', self._on_file_changed)
        self.inv_path.trace_add('write',   self._on_file_changed)

    # ── layout ──────────────────────────────────────────────────────────────

    def _center(self):
        self.update_idletasks()
        w, h = 1020, 820
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _card(self, parent, title):
        frame = tk.Frame(parent, bg=CARD_BG,
                         highlightbackground=BORDER, highlightthickness=1)
        tk.Label(frame, text=title,
                 bg=PRIMARY, fg='white',
                 font=('Segoe UI', 10, 'bold'),
                 padx=12, pady=5).grid(row=0, column=0, columnspan=10, sticky='ew')
        frame.grid_columnconfigure(0, weight=0)
        frame.grid_columnconfigure(1, weight=1)
        frame.grid_columnconfigure(2, weight=1)
        return frame

    def _build_ui(self):
        # ── header ──────────────────────────────────────────────────────────
        header = tk.Frame(self, bg=PRIMARY, height=72)
        header.pack(fill='x')
        header.pack_propagate(False)

        logo_path = resource_path(os.path.join('File_template', 'Bluecircle.png'))
        try:
            img = Image.open(logo_path).convert('RGBA')
            img.thumbnail((52, 52), Image.LANCZOS)
            self._logo_img = ImageTk.PhotoImage(img)
            tk.Label(header, image=self._logo_img, bg=PRIMARY).pack(side='left', padx=16, pady=10)
        except Exception:
            pass

        title_frame = tk.Frame(header, bg=PRIMARY)
        title_frame.pack(side='left', pady=8)
        tk.Label(title_frame, text=APP_TITLE,
                 font=('Segoe UI', 16, 'bold'),
                 fg='white', bg=PRIMARY).pack(anchor='w')
        tk.Label(title_frame,
                 text="Phân tích sức bán & đề xuất luân chuyển hàng hoá bán lẻ",
                 font=('Segoe UI', 9), fg='#BBDEFB', bg=PRIMARY).pack(anchor='w')

        # ── main body ───────────────────────────────────────────────────────
        body = tk.Frame(self, bg=BG)
        body.pack(fill='both', expand=True, padx=20, pady=12)

        self._build_import_card(body)
        self._build_store_card(body)
        self._build_settings_card(body)
        self._build_action_bar(body)

    # ── import card ──────────────────────────────────────────────────────────

    def _build_import_card(self, parent):
        card = self._card(parent, "Nhập dữ liệu")
        card.pack(fill='x', pady=(0, 8))

        row = tk.Frame(card, bg=CARD_BG)
        row.grid(row=1, column=0, columnspan=3, sticky='ew', padx=14, pady=(8, 4))
        for i in (1, 4):
            row.grid_columnconfigure(i, weight=1)

        # File bán hàng
        tk.Label(row, text="File bán hàng:", bg=CARD_BG, fg=TEXT_DARK,
                 font=('Segoe UI', 10)).grid(row=0, column=0, sticky='w', padx=(0, 6))
        tk.Entry(row, textvariable=self.sales_path, width=32,
                 font=('Segoe UI', 9), fg=TEXT_MID, relief='flat', bd=0,
                 highlightbackground=BORDER, highlightthickness=1
                 ).grid(row=0, column=1, sticky='ew', padx=(0, 4))
        tk.Button(row, text="Chọn file",
                  command=lambda: self._browse(self.sales_path),
                  bg=PRIMARY_L, fg='white', activebackground=ACCENT,
                  font=('Segoe UI', 9), relief='flat', cursor='hand2',
                  padx=10, pady=3
                  ).grid(row=0, column=2, padx=(0, 20))

        tk.Frame(row, bg=BORDER, width=1, height=28
                 ).grid(row=0, column=3, sticky='ns', padx=4)

        # File tồn kho
        tk.Label(row, text="File tồn kho:", bg=CARD_BG, fg=TEXT_DARK,
                 font=('Segoe UI', 10)).grid(row=0, column=4, sticky='w', padx=(20, 6))
        tk.Entry(row, textvariable=self.inv_path, width=32,
                 font=('Segoe UI', 9), fg=TEXT_MID, relief='flat', bd=0,
                 highlightbackground=BORDER, highlightthickness=1
                 ).grid(row=0, column=5, sticky='ew', padx=(0, 4))
        tk.Button(row, text="Chọn file",
                  command=lambda: self._browse(self.inv_path),
                  bg=PRIMARY_L, fg='white', activebackground=ACCENT,
                  font=('Segoe UI', 9), relief='flat', cursor='hand2',
                  padx=10, pady=3
                  ).grid(row=0, column=6)

        # ── Đọc File button row ──────────────────────────────────────────────
        btn_row = tk.Frame(card, bg=CARD_BG)
        btn_row.grid(row=2, column=0, columnspan=3, sticky='ew', padx=14, pady=(4, 10))

        self.btn_read = tk.Button(
            btn_row, text="  Đọc File  ",
            command=self._on_read_file_click,
            bg=ACCENT, fg='white', activebackground=PRIMARY,
            font=('Segoe UI', 10, 'bold'),
            relief='flat', cursor='hand2', padx=16, pady=6,
            state='disabled',
        )
        self.btn_read.pack(side='left')

        self.read_status = tk.Label(
            btn_row, text="",
            bg=CARD_BG, fg=TEXT_LIGHT, font=('Segoe UI', 9),
        )
        self.read_status.pack(side='left', padx=12)

    # ── store selection card ─────────────────────────────────────────────────

    def _build_store_card(self, parent):
        card = self._card(parent, "Lựa chọn cửa hàng phân tích")
        card.pack(fill='x', pady=(0, 8))
        card.grid_columnconfigure(0, weight=1)
        card.grid_columnconfigure(1, weight=1)

        # ── Kho Cửa Hàng (stores with recent sales) ──────────────────────────
        ph = tk.Frame(card, bg=CARD_BG)
        ph.grid(row=1, column=0, sticky='nsew', padx=(14, 6), pady=(8, 10))

        ph_top = tk.Frame(ph, bg=CARD_BG)
        ph_top.pack(fill='x')
        tk.Label(ph_top, text="Kho Cửa Hàng",
                 bg=CARD_BG, fg=TEXT_DARK,
                 font=('Segoe UI', 9, 'bold')).pack(side='left')
        self.lbl_physical_count = tk.Label(ph_top, text="(chưa tải file)",
                                            bg=CARD_BG, fg=TEXT_LIGHT,
                                            font=('Segoe UI', 8))
        self.lbl_physical_count.pack(side='left', padx=6)

        ph_btns = tk.Frame(ph, bg=CARD_BG)
        ph_btns.pack(fill='x', pady=(3, 2))
        tk.Button(ph_btns, text="Chọn tất cả", font=('Segoe UI', 8),
                  bg='#E3F2FD', fg=PRIMARY, relief='flat', cursor='hand2', padx=6, pady=1,
                  command=lambda: self.lb_physical.select_all()
                  ).pack(side='left', padx=(0, 4))
        tk.Button(ph_btns, text="Bỏ chọn tất cả", font=('Segoe UI', 8),
                  bg='#FFF3E0', fg=WARNING, relief='flat', cursor='hand2', padx=6, pady=1,
                  command=lambda: self.lb_physical.deselect_all()
                  ).pack(side='left')

        self._ph_search = tk.StringVar()
        ph_search_entry = tk.Entry(
            ph, textvariable=self._ph_search,
            font=('Segoe UI', 9), fg=TEXT_MID,
            relief='flat', bd=0,
            highlightbackground=BORDER, highlightthickness=1,
        )
        ph_search_entry.pack(fill='x', pady=(2, 4))
        ph_search_entry.insert(0, "Tìm kiếm kho...")
        ph_search_entry.config(fg=TEXT_LIGHT)
        ph_search_entry.bind('<FocusIn>',  lambda e: self._search_focus_in(ph_search_entry,  self._ph_search))
        ph_search_entry.bind('<FocusOut>', lambda e: self._search_focus_out(ph_search_entry, self._ph_search, "Tìm kiếm kho..."))
        self._ph_search.trace_add('write', lambda *_: self.lb_physical.filter(
            '' if self._ph_search.get() == "Tìm kiếm kho..." else self._ph_search.get()
        ))

        self.lb_physical = CheckListbox(ph, height=130)
        self.lb_physical.pack(fill='both', expand=True)

        # ── Các kho khác (Shopee + no-recent-sales + specific warehouses) ────
        sp = tk.Frame(card, bg=CARD_BG)
        sp.grid(row=1, column=1, sticky='nsew', padx=(6, 14), pady=(8, 10))

        sp_top = tk.Frame(sp, bg=CARD_BG)
        sp_top.pack(fill='x')
        tk.Label(sp_top, text="Các kho khác",
                 bg=CARD_BG, fg=TEXT_DARK,
                 font=('Segoe UI', 9, 'bold')).pack(side='left')
        self.lbl_shopee_count = tk.Label(sp_top, text="(chưa tải file)",
                                          bg=CARD_BG, fg=TEXT_LIGHT,
                                          font=('Segoe UI', 8))
        self.lbl_shopee_count.pack(side='left', padx=6)

        sp_btns = tk.Frame(sp, bg=CARD_BG)
        sp_btns.pack(fill='x', pady=(3, 2))
        tk.Button(sp_btns, text="Chọn tất cả", font=('Segoe UI', 8),
                  bg='#E3F2FD', fg=PRIMARY, relief='flat', cursor='hand2', padx=6, pady=1,
                  command=lambda: self.lb_shopee.select_all()
                  ).pack(side='left', padx=(0, 4))
        tk.Button(sp_btns, text="Bỏ chọn tất cả", font=('Segoe UI', 8),
                  bg='#FFF3E0', fg=WARNING, relief='flat', cursor='hand2', padx=6, pady=1,
                  command=lambda: self.lb_shopee.deselect_all()
                  ).pack(side='left')

        self._sp_search = tk.StringVar()
        sp_search_entry = tk.Entry(
            sp, textvariable=self._sp_search,
            font=('Segoe UI', 9), fg=TEXT_MID,
            relief='flat', bd=0,
            highlightbackground=BORDER, highlightthickness=1,
        )
        sp_search_entry.pack(fill='x', pady=(2, 4))
        sp_search_entry.insert(0, "Tìm kiếm kho...")
        sp_search_entry.config(fg=TEXT_LIGHT)
        sp_search_entry.bind('<FocusIn>',  lambda e: self._search_focus_in(sp_search_entry,  self._sp_search))
        sp_search_entry.bind('<FocusOut>', lambda e: self._search_focus_out(sp_search_entry, self._sp_search, "Tìm kiếm kho..."))
        self._sp_search.trace_add('write', lambda *_: self.lb_shopee.filter(
            '' if self._sp_search.get() == "Tìm kiếm kho..." else self._sp_search.get()
        ))

        self.lb_shopee = CheckListbox(sp, height=130)
        self.lb_shopee.pack(fill='both', expand=True)

    # ── settings card ────────────────────────────────────────────────────────

    def _build_settings_card(self, parent):
        card = self._card(parent, "Tham số & Quy tắc tính toán")
        card.pack(fill='x', pady=(0, 8))

        # Row 1: target months + slow pct
        tk.Label(card, text="Mức tồn mục tiêu (tháng):",
                 bg=CARD_BG, fg=TEXT_MID,
                 font=('Segoe UI', 10)).grid(row=1, column=0, sticky='w',
                                              padx=(14, 6), pady=(8, 4))
        ttk.Spinbox(card, from_=0.5, to=6.0, increment=0.5,
                    textvariable=self.target_months, width=7,
                    font=('Segoe UI', 10)
                    ).grid(row=1, column=1, sticky='w', pady=(8, 4))
        tk.Label(card, text="(Fill hàng về cửa hàng để đủ X tháng sức bán)",
                 bg=CARD_BG, fg=TEXT_LIGHT,
                 font=('Segoe UI', 9)).grid(row=1, column=2, sticky='w', padx=10, pady=(8, 4))

        tk.Label(card, text="Ngưỡng bán chậm (% sức bán TB mạng lưới):",
                 bg=CARD_BG, fg=TEXT_MID,
                 font=('Segoe UI', 10)).grid(row=2, column=0, sticky='w',
                                              padx=(14, 6), pady=4)
        ttk.Spinbox(card, from_=0.0, to=100.0, increment=5.0,
                    textvariable=self.slow_pct, width=7,
                    font=('Segoe UI', 10)
                    ).grid(row=2, column=1, sticky='w', pady=4)
        tk.Label(card, text="(Dưới ngưỡng này = bán chậm, đề xuất luân chuyển đi)",
                 bg=CARD_BG, fg=TEXT_LIGHT,
                 font=('Segoe UI', 9)).grid(row=2, column=2, sticky='w', padx=10, pady=4)

        # Row 3: min fill avg threshold
        tk.Label(card, text="Ngưỡng fill tối thiểu (TB/tháng):",
                 bg=CARD_BG, fg=TEXT_MID,
                 font=('Segoe UI', 10)).grid(row=3, column=0, sticky='w',
                                              padx=(14, 6), pady=4)
        ttk.Spinbox(card, from_=0.0, to=20.0, increment=0.1,
                    textvariable=self.min_fill_avg, width=7,
                    font=('Segoe UI', 10), format='%.1f'
                    ).grid(row=3, column=1, sticky='w', pady=4)
        tk.Label(card, text="(≤ X SP/tháng: bỏ qua, không fill về cửa hàng)",
                 bg=CARD_BG, fg=TEXT_LIGHT,
                 font=('Segoe UI', 9)).grid(row=3, column=2, sticky='w', padx=10, pady=4)

        # Row 4: HCM <> HN rule
        rule_frame = tk.Frame(card, bg=CARD_BG)
        rule_frame.grid(row=4, column=0, columnspan=3, sticky='w',
                        padx=(14, 14), pady=(4, 10))
        tk.Checkbutton(
            rule_frame,
            text="Cho phép luân chuyển hàng hoá giữa  HCM  ↔  HN",
            variable=self.allow_hcm_hn,
            bg=CARD_BG, fg=TEXT_DARK,
            font=('Segoe UI', 10),
            activebackground=CARD_BG,
            selectcolor='white',
        ).pack(side='left')
        tk.Label(rule_frame,
                 text="  (nếu không chọn: chỉ luân chuyển trong cùng khu vực)",
                 bg=CARD_BG, fg=TEXT_LIGHT,
                 font=('Segoe UI', 9)).pack(side='left')

    # ── action bar ───────────────────────────────────────────────────────────

    def _build_action_bar(self, parent):
        btn_bar = tk.Frame(parent, bg=BG)
        btn_bar.pack(fill='x', pady=(0, 8))

        self.btn_export = tk.Button(
            btn_bar, text="  Xuất Excel  ",
            command=self._export,
            bg=SUCCESS, fg='white', activebackground='#388E3C',
            font=('Segoe UI', 11, 'bold'),
            relief='flat', cursor='hand2', padx=18, pady=8,
            state='disabled',
        )
        self.btn_export.pack(side='left')

        self.progress = ttk.Progressbar(btn_bar, mode='indeterminate', length=180)
        self.progress.pack(side='left', padx=20)

        tk.Label(btn_bar, textvariable=self.status_text,
                 bg=BG, fg=TEXT_MID, font=('Segoe UI', 9)).pack(side='left')

    # ── file browse ──────────────────────────────────────────────────────────

    def _browse(self, var):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            var.set(path)

    # ── search placeholder helpers ────────────────────────────────────────────

    def _search_focus_in(self, entry, var):
        if entry.cget('fg') == TEXT_LIGHT:
            entry.delete(0, 'end')
            entry.config(fg=TEXT_DARK)

    def _search_focus_out(self, entry, var, placeholder):
        if not var.get():
            entry.insert(0, placeholder)
            entry.config(fg=TEXT_LIGHT)

    # ── file changed → enable Đọc File button ────────────────────────────────

    def _on_file_changed(self, *args):
        s = self.sales_path.get()
        i = self.inv_path.get()
        if s and i and os.path.exists(s) and os.path.exists(i):
            self.btn_read.config(state='normal')
            self.read_status.config(text="Sẵn sàng đọc file")
        else:
            self.btn_read.config(state='disabled')

    def _on_read_file_click(self):
        self.btn_read.config(state='disabled')
        self.btn_export.config(state='disabled')
        self.read_status.config(text="Đang đọc file...")
        self.lbl_physical_count.config(text="(đang tải...)")
        self.lbl_shopee_count.config(text="(đang tải...)")
        threading.Thread(target=self._load_stores_thread, daemon=True).start()

    def _load_stores_thread(self):
        try:
            df = pd.read_excel(self.sales_path.get())
            df_inv = pd.read_excel(self.inv_path.get())
            df.columns = df.columns.str.strip()
            self.df_sales_raw = df
            self.df_inv_raw   = df_inv

            retail = df[df['Sale Team'].str.strip() == 'Bán Lẻ'].copy()
            retail['Date'] = pd.to_datetime(retail['Date'], errors='coerce')
            all_stores = sorted(
                retail['location Name'].dropna().astype(str).unique().tolist()
            )

            # Stores that have sales in the last 3 months
            stores_with_recent: set = set()
            if not retail.empty:
                max_date = retail['Date'].max()
                min_3m   = max_date - pd.DateOffset(months=3)
                recent   = retail[retail['Date'] >= min_3m]
                stores_with_recent = set(
                    recent['location Name'].dropna().astype(str).unique()
                )

            # Keywords that always go to "Các kho khác" regardless of sales
            OTHER_KW = ('shopee', 'droppi', 'xe thuê', 'kho kg')

            def is_other(name: str) -> bool:
                n = name.lower()
                return any(kw in n for kw in OTHER_KW) or name not in stores_with_recent

            physical = [s for s in all_stores if not is_other(s)]
            shopee   = [s for s in all_stores if is_other(s)]

            self.after(0, self._populate_store_lists, physical, shopee)
        except Exception:
            self.after(0, self._on_store_load_error)

    def _populate_store_lists(self, physical, shopee):
        self._physical_stores = physical
        self._shopee_stores   = shopee

        self.lb_physical.set_items(physical, select_all=True)
        self.lb_shopee.set_items(shopee, select_all=True)

        self.lbl_physical_count.config(text=f"({len(physical)} cửa hàng)")
        self.lbl_shopee_count.config(
            text=f"({len(shopee)} kho)" if shopee else "(không có)"
        )
        self.btn_read.config(state='normal')
        self.btn_export.config(state='normal')
        self.read_status.config(text=f"Đã tải  •  {len(physical)} kho cửa hàng  •  {len(shopee)} kho khác")
        self.status_text.set("Chọn kho và bấm Xuất Excel")

    def _on_store_load_error(self):
        self.lbl_physical_count.config(text="(lỗi đọc file)")
        self.lbl_shopee_count.config(text="(lỗi đọc file)")
        self.btn_read.config(state='normal')
        self.read_status.config(text="Lỗi đọc file — kiểm tra lại định dạng")

    # ── helpers ──────────────────────────────────────────────────────────────

    def _get_selected_stores(self):
        """Return selected store names (None = include all when no list loaded)."""
        if not self._physical_stores and not self._shopee_stores:
            return None  # files not loaded yet; calculate() will use all
        return self.lb_physical.get_selected() + self.lb_shopee.get_selected()

    # ── export (phân tích + xuất trong 1 bước) ───────────────────────────────

    def _export(self):
        selected = self._get_selected_stores()
        if selected is not None and len(selected) == 0:
            messagebox.showwarning("Chưa chọn kho",
                                   "Vui lòng chọn ít nhất một kho để phân tích.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"LuanChuyenHangHoa_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        )
        if not save_path:
            return

        self.btn_export.config(state='disabled')
        self.status_text.set("Đang phân tích & xuất file...")
        self.progress.start(12)
        threading.Thread(
            target=self._export_worker,
            args=(save_path, selected),
            daemon=True,
        ).start()

    def _export_worker(self, save_path, selected_stores):
        try:
            df_sales = (self.df_sales_raw if self.df_sales_raw is not None
                        else pd.read_excel(self.sales_path.get()))
            df_inv   = (self.df_inv_raw   if self.df_inv_raw   is not None
                        else pd.read_excel(self.inv_path.get()))

            df_fill, df_transfer, df_summary, warnings = calculate(
                df_sales, df_inv,
                target_months   = self.target_months.get(),
                slow_pct        = self.slow_pct.get(),
                selected_stores = selected_stores,
                allow_hcm_hn    = self.allow_hcm_hn.get(),
                min_fill_avg    = self.min_fill_avg.get(),
            )
            export_excel(df_fill, df_transfer, df_summary, save_path)
            self.after(0, self._on_export_done, save_path, warnings, df_fill, df_transfer)
        except Exception as e:
            self.after(0, self._on_export_error, str(e))

    def _on_export_done(self, save_path, warnings, df_fill, df_transfer):
        self.progress.stop()
        self.btn_export.config(state='normal')
        if warnings:
            messagebox.showwarning("Cảnh báo", "\n".join(warnings))
        total_fill  = int(df_fill['Fill thực tế'].sum())           if len(df_fill)     else 0
        total_trans = int(df_transfer['Đề xuất luân chuyển'].sum()) if len(df_transfer) else 0
        self.status_text.set(
            f"Hoàn thành  •  Fill: {total_fill:,} SP  •  Luân chuyển: {total_trans:,} SP"
        )
        messagebox.showinfo(
            "Xuất file thành công",
            f"Đã lưu:\n{save_path}\n\n"
            f"• Fill từ kho:            {total_fill:,} SP\n"
            f"• Luân chuyển cửa hàng:  {total_trans:,} SP",
        )

    def _on_export_error(self, msg):
        self.progress.stop()
        self.btn_export.config(state='normal')
        self.status_text.set("Lỗi!")
        messagebox.showerror("Lỗi", msg)


# ─── Entry ────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    # High-DPI awareness on Windows (auto-scales to monitor DPI)
    if sys.platform == 'win32':
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(2)   # Per-monitor DPI aware v1
        except Exception:
            try:
                windll.user32.SetProcessDPIAware()    # Fallback: system DPI aware
            except Exception:
                pass
    app = App()
    app.mainloop()
