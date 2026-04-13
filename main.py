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

    prod_names = (
        retail_sales[['SKU', 'Product Item']]
        .drop_duplicates('SKU')
        .rename(columns={'Product Item': 'ProductName'})
    )

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

    df_fill = fill[fill['FinalFill'] > 0].copy()
    df_fill = df_fill[[
        'Store', 'SKU', 'ProductName',
        'Sold3M', 'AvgMonthly', 'WeeksOfStock',
        'StoreQty', 'TargetStock', 'SuggestedFill',
        'KhoQty', 'FinalFill',
    ]].rename(columns={
        'Store':         'Cửa hàng',
        'SKU':           'Mã hàng',
        'ProductName':   'Tên sản phẩm',
        'Sold3M':        'Đã bán 3T',
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

    _empty_transfer = pd.DataFrame(columns=[
        'Mã hàng', 'Tên sản phẩm',
        'Cửa hàng gửi', 'Tồn gửi', 'Sức bán gửi TB/tháng',
        'Cửa hàng nhận', 'Tồn nhận', 'Sức bán nhận TB/tháng',
        'Đề xuất luân chuyển',
    ])

    if not grid_rows:
        return df_fill, _empty_transfer, warnings

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
    else:
        df_transfer = _empty_transfer

    return df_fill, df_transfer, warnings


def export_excel(df_fill: pd.DataFrame, df_transfer: pd.DataFrame, save_path: str):
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
                        highlight_col=None, hl_fmts=None):
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
                        if col == highlight_col and hl_fmts:
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
            sum_cols      = ['Đã bán 3T', 'Tồn cửa hàng', 'Mức tồn mục tiêu',
                             'Đề xuất fill', 'Fill thực tế'],
            highlight_col = 'Fill thực tế',
            hl_fmts       = hl_fill,
        )
        write_sheet(
            df_transfer, 'Luân chuyển cửa hàng',
            title_text    = 'PHÂN TÍCH LUÂN CHUYỂN HÀNG HOÁ GIỮA CÁC CỬA HÀNG',
            group_col     = 'Cửa hàng gửi',
            sum_cols      = ['Tồn gửi', 'Đề xuất luân chuyển'],
            highlight_col = 'Đề xuất luân chuyển',
            hl_fmts       = hl_xfer,
        )


# ─── GUI ──────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.resizable(True, True)
        self.minsize(900, 700)
        self.configure(bg=BG)

        # state
        self.sales_path    = tk.StringVar()
        self.inv_path      = tk.StringVar()
        self.target_months = tk.DoubleVar(value=1.5)
        self.slow_pct      = tk.DoubleVar(value=50.0)
        self.min_fill_avg  = tk.DoubleVar(value=0.4)
        self.allow_hcm_hn  = tk.BooleanVar(value=False)
        self.status_text   = tk.StringVar(value="Chưa phân tích")

        self._physical_stores: list = []
        self._shopee_stores:   list = []
        self.df_sales_raw = None
        self.df_inv_raw   = None
        self.df_fill      = None
        self.df_transfer  = None

        self._build_ui()
        self._center()

        # Auto-load store list when both files selected
        self.sales_path.trace_add('write', self._on_file_changed)
        self.inv_path.trace_add('write',   self._on_file_changed)

    # ── layout ──────────────────────────────────────────────────────────────

    def _center(self):
        self.update_idletasks()
        w, h = 1020, 840
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
        self._build_results(body)

    # ── import card (compact, both files on one row) ─────────────────────────

    def _build_import_card(self, parent):
        card = self._card(parent, "Nhập dữ liệu")
        card.pack(fill='x', pady=(0, 8))

        row = tk.Frame(card, bg=CARD_BG)
        row.grid(row=1, column=0, columnspan=3, sticky='ew', padx=14, pady=8)
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

        # separator line
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

    # ── store selection card ─────────────────────────────────────────────────

    def _build_store_card(self, parent):
        card = self._card(parent, "Lựa chọn cửa hàng phân tích")
        card.pack(fill='x', pady=(0, 8))
        card.grid_columnconfigure(0, weight=1)
        card.grid_columnconfigure(1, weight=1)

        # ── Physical stores ──────────────────────────────────────────────────
        ph = tk.Frame(card, bg=CARD_BG)
        ph.grid(row=1, column=0, sticky='nsew', padx=(14, 6), pady=(8, 10))

        ph_top = tk.Frame(ph, bg=CARD_BG)
        ph_top.pack(fill='x')
        tk.Label(ph_top, text="Cửa hàng vật lý",
                 bg=CARD_BG, fg=TEXT_DARK,
                 font=('Segoe UI', 9, 'bold')).pack(side='left')
        self.lbl_physical_count = tk.Label(ph_top, text="(chưa tải file)",
                                            bg=CARD_BG, fg=TEXT_LIGHT,
                                            font=('Segoe UI', 8))
        self.lbl_physical_count.pack(side='left', padx=6)

        ph_btns = tk.Frame(ph, bg=CARD_BG)
        ph_btns.pack(fill='x', pady=(3, 3))
        tk.Button(ph_btns, text="Chọn tất cả", font=('Segoe UI', 8),
                  bg='#E3F2FD', fg=PRIMARY, relief='flat', cursor='hand2', padx=6, pady=1,
                  command=lambda: self.lb_physical.select_set(0, tk.END)
                  ).pack(side='left', padx=(0, 4))
        tk.Button(ph_btns, text="Bỏ chọn tất cả", font=('Segoe UI', 8),
                  bg='#FFF3E0', fg=WARNING, relief='flat', cursor='hand2', padx=6, pady=1,
                  command=lambda: self.lb_physical.select_clear(0, tk.END)
                  ).pack(side='left')

        ph_frame = tk.Frame(ph, bg=CARD_BG,
                            highlightbackground=BORDER, highlightthickness=1)
        ph_frame.pack(fill='both', expand=True)
        ph_vsb = ttk.Scrollbar(ph_frame, orient='vertical')
        ph_vsb.pack(side='right', fill='y')
        self.lb_physical = tk.Listbox(
            ph_frame, selectmode=tk.MULTIPLE,
            yscrollcommand=ph_vsb.set,
            height=6, font=('Segoe UI', 9),
            bg='#FAFAFA', fg=TEXT_DARK,
            selectbackground=PRIMARY_L, selectforeground='white',
            relief='flat', bd=0,
            activestyle='none',
        )
        self.lb_physical.pack(fill='both', expand=True)
        ph_vsb.config(command=self.lb_physical.yview)

        # ── Shopee stores ────────────────────────────────────────────────────
        sp = tk.Frame(card, bg=CARD_BG)
        sp.grid(row=1, column=1, sticky='nsew', padx=(6, 14), pady=(8, 10))

        sp_top = tk.Frame(sp, bg=CARD_BG)
        sp_top.pack(fill='x')
        tk.Label(sp_top, text="Kênh Shopee",
                 bg=CARD_BG, fg=TEXT_DARK,
                 font=('Segoe UI', 9, 'bold')).pack(side='left')
        self.lbl_shopee_count = tk.Label(sp_top, text="(chưa tải file)",
                                          bg=CARD_BG, fg=TEXT_LIGHT,
                                          font=('Segoe UI', 8))
        self.lbl_shopee_count.pack(side='left', padx=6)

        sp_btns = tk.Frame(sp, bg=CARD_BG)
        sp_btns.pack(fill='x', pady=(3, 3))
        tk.Button(sp_btns, text="Chọn tất cả", font=('Segoe UI', 8),
                  bg='#E3F2FD', fg=PRIMARY, relief='flat', cursor='hand2', padx=6, pady=1,
                  command=lambda: self.lb_shopee.select_set(0, tk.END)
                  ).pack(side='left', padx=(0, 4))
        tk.Button(sp_btns, text="Bỏ chọn tất cả", font=('Segoe UI', 8),
                  bg='#FFF3E0', fg=WARNING, relief='flat', cursor='hand2', padx=6, pady=1,
                  command=lambda: self.lb_shopee.select_clear(0, tk.END)
                  ).pack(side='left')

        sp_frame = tk.Frame(sp, bg=CARD_BG,
                            highlightbackground=BORDER, highlightthickness=1)
        sp_frame.pack(fill='both', expand=True)
        sp_vsb = ttk.Scrollbar(sp_frame, orient='vertical')
        sp_vsb.pack(side='right', fill='y')
        self.lb_shopee = tk.Listbox(
            sp_frame, selectmode=tk.MULTIPLE,
            yscrollcommand=sp_vsb.set,
            height=6, font=('Segoe UI', 9),
            bg='#FAFAFA', fg=TEXT_DARK,
            selectbackground='#E64A19', selectforeground='white',
            relief='flat', bd=0,
            activestyle='none',
        )
        self.lb_shopee.pack(fill='both', expand=True)
        sp_vsb.config(command=self.lb_shopee.yview)

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

        self.btn_analyze = tk.Button(
            btn_bar, text="  Phân Tích  ",
            command=self._run_analysis,
            bg=PRIMARY, fg='white', activebackground=PRIMARY_L,
            font=('Segoe UI', 11, 'bold'),
            relief='flat', cursor='hand2', padx=18, pady=8,
        )
        self.btn_analyze.pack(side='left', padx=(0, 10))

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

    # ── results ──────────────────────────────────────────────────────────────

    def _build_results(self, parent):
        result_card = tk.Frame(parent, bg=CARD_BG,
                               highlightbackground=BORDER, highlightthickness=1)
        result_card.pack(fill='both', expand=True)

        style = ttk.Style()
        style.configure('TNotebook', background=CARD_BG, borderwidth=0)
        style.configure('TNotebook.Tab',
                        font=('Segoe UI', 10, 'bold'), padding=[14, 6])

        self.notebook = ttk.Notebook(result_card)
        self.notebook.pack(fill='both', expand=True, padx=2, pady=2)

        self.tab_fill     = self._make_tab("Fill từ kho")
        self.tab_transfer = self._make_tab("Luân chuyển cửa hàng")

        self.lbl_fill_sum  = tk.StringVar(value="")
        self.lbl_trans_sum = tk.StringVar(value="")
        tk.Label(self.tab_fill, textvariable=self.lbl_fill_sum,
                 bg=CARD_BG, fg=TEXT_MID,
                 font=('Segoe UI', 9)).pack(anchor='e', padx=8)
        tk.Label(self.tab_transfer, textvariable=self.lbl_trans_sum,
                 bg=CARD_BG, fg=TEXT_MID,
                 font=('Segoe UI', 9)).pack(anchor='e', padx=8)

        self.tree_fill     = self._make_tree(self.tab_fill)
        self.tree_transfer = self._make_tree(self.tab_transfer)

    def _make_tab(self, label):
        frame = tk.Frame(self.notebook, bg=CARD_BG)
        self.notebook.add(frame, text=f"  {label}  ")
        return frame

    def _make_tree(self, parent):
        frame = tk.Frame(parent, bg=CARD_BG)
        frame.pack(fill='both', expand=True, padx=4, pady=(0, 4))

        vsb = ttk.Scrollbar(frame, orient='vertical')
        hsb = ttk.Scrollbar(frame, orient='horizontal')
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')

        tree = ttk.Treeview(frame,
                            yscrollcommand=vsb.set,
                            xscrollcommand=hsb.set,
                            show='headings')
        tree.pack(fill='both', expand=True)
        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)

        style = ttk.Style()
        style.configure('Treeview', rowheight=22, font=('Segoe UI', 9))
        style.configure('Treeview.Heading',
                        font=('Segoe UI', 9, 'bold'),
                        background='#1565C0', foreground='white')
        tree.tag_configure('odd',  background='#F5F7FA')
        tree.tag_configure('even', background='#FFFFFF')
        return tree

    # ── file browse ──────────────────────────────────────────────────────────

    def _browse(self, var):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            var.set(path)

    # ── auto-load stores when files change ───────────────────────────────────

    def _on_file_changed(self, *args):
        s = self.sales_path.get()
        i = self.inv_path.get()
        if s and i and os.path.exists(s) and os.path.exists(i):
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

            retail = df[df['Sale Team'].str.strip() == 'Bán Lẻ']
            stores = sorted(retail['location Name'].dropna().astype(str).unique().tolist())

            physical = [s for s in stores if 'shopee' not in s.lower()]
            shopee   = [s for s in stores if 'shopee' in s.lower()]

            self.after(0, self._populate_store_lists, physical, shopee)
        except Exception:
            self.after(0, self._on_store_load_error)

    def _populate_store_lists(self, physical, shopee):
        self._physical_stores = physical
        self._shopee_stores   = shopee

        self.lb_physical.delete(0, tk.END)
        for s in physical:
            self.lb_physical.insert(tk.END, s)
        self.lb_physical.select_set(0, tk.END)

        self.lb_shopee.delete(0, tk.END)
        for s in shopee:
            self.lb_shopee.insert(tk.END, s)
        self.lb_shopee.select_set(0, tk.END)

        self.lbl_physical_count.config(text=f"({len(physical)} cửa hàng)")
        self.lbl_shopee_count.config(
            text=f"({len(shopee)} kênh)" if shopee else "(không có)"
        )

    def _on_store_load_error(self):
        self.lbl_physical_count.config(text="(lỗi đọc file)")
        self.lbl_shopee_count.config(text="(lỗi đọc file)")

    # ── helpers ──────────────────────────────────────────────────────────────

    def _get_selected_stores(self):
        """Return selected store names (None = include all when no list loaded)."""
        if not self._physical_stores and not self._shopee_stores:
            return None  # files not loaded yet; calculate() will use all
        physical_sel = [self._physical_stores[i]
                        for i in self.lb_physical.curselection()]
        shopee_sel   = [self._shopee_stores[i]
                        for i in self.lb_shopee.curselection()]
        return physical_sel + shopee_sel

    # ── analysis ─────────────────────────────────────────────────────────────

    def _run_analysis(self):
        if not self.sales_path.get():
            messagebox.showwarning("Thiếu file", "Vui lòng chọn file dữ liệu bán hàng.")
            return
        if not self.inv_path.get():
            messagebox.showwarning("Thiếu file", "Vui lòng chọn file tồn kho.")
            return

        selected = self._get_selected_stores()
        if selected is not None and len(selected) == 0:
            messagebox.showwarning("Chưa chọn cửa hàng",
                                   "Vui lòng chọn ít nhất một cửa hàng để phân tích.")
            return

        self.btn_analyze.config(state='disabled')
        self.btn_export.config(state='disabled')
        self.status_text.set("Đang phân tích...")
        self.progress.start(12)

        threading.Thread(target=self._analysis_worker, daemon=True).start()

    def _analysis_worker(self):
        try:
            df_sales = (self.df_sales_raw if self.df_sales_raw is not None
                        else pd.read_excel(self.sales_path.get()))
            df_inv   = (self.df_inv_raw   if self.df_inv_raw   is not None
                        else pd.read_excel(self.inv_path.get()))

            df_fill, df_transfer, warnings = calculate(
                df_sales, df_inv,
                target_months  = self.target_months.get(),
                slow_pct       = self.slow_pct.get(),
                selected_stores= self._get_selected_stores(),
                allow_hcm_hn   = self.allow_hcm_hn.get(),
                min_fill_avg   = self.min_fill_avg.get(),
            )

            self.df_fill     = df_fill
            self.df_transfer = df_transfer
            self.after(0, self._on_analysis_done, warnings)
        except Exception as e:
            self.after(0, self._on_analysis_error, str(e))

    def _on_analysis_done(self, warnings):
        self.progress.stop()
        self.btn_analyze.config(state='normal')

        if warnings:
            messagebox.showwarning("Cảnh báo", "\n".join(warnings))

        self._populate_tree(self.tree_fill,     self.df_fill)
        self._populate_tree(self.tree_transfer, self.df_transfer)

        fill_rows  = len(self.df_fill)
        trans_rows = len(self.df_transfer)
        total_fill  = int(self.df_fill['Fill thực tế'].sum())      if fill_rows  else 0
        total_trans = int(self.df_transfer['Đề xuất luân chuyển'].sum()) if trans_rows else 0

        self.lbl_fill_sum.set(f"{fill_rows} dòng  |  Tổng fill: {total_fill:,} SP")
        self.lbl_trans_sum.set(f"{trans_rows} dòng  |  Tổng luân chuyển: {total_trans:,} SP")
        self.status_text.set(
            f"Hoàn thành  •  Fill: {total_fill:,} SP  •  Luân chuyển: {total_trans:,} SP"
        )
        self.btn_export.config(state='normal')

    def _on_analysis_error(self, msg):
        self.progress.stop()
        self.btn_analyze.config(state='normal')
        self.status_text.set("Lỗi!")
        messagebox.showerror("Lỗi phân tích", msg)

    def _populate_tree(self, tree, df):
        tree.delete(*tree.get_children())
        if df is None or df.empty:
            tree['columns'] = ('empty',)
            tree.heading('empty', text='Không có dữ liệu')
            tree.column('empty', width=300)
            return

        cols = list(df.columns)
        tree['columns'] = cols
        for col in cols:
            tree.heading(col, text=col, anchor='center')
            w = max(len(col) * 9, 80)
            tree.column(col, width=w, anchor='center', minwidth=60)

        for i, row in enumerate(df.itertuples(index=False)):
            tag = 'odd' if i % 2 else 'even'
            values = []
            for v in row:
                if isinstance(v, float):
                    values.append(f"{v:,.2f}")
                elif isinstance(v, (int, np.integer)):
                    values.append(f"{v:,}")
                else:
                    values.append(str(v) if v is not None else '')
            tree.insert('', 'end', values=values, tags=(tag,))

    # ── export ───────────────────────────────────────────────────────────────

    def _export(self):
        if self.df_fill is None:
            return
        save_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"LuanChuyenHangHoa_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        )
        if not save_path:
            return
        try:
            export_excel(self.df_fill, self.df_transfer, save_path)
            messagebox.showinfo(
                "Thành công",
                f"Đã xuất file:\n{save_path}\n\n"
                "• Sheet 'Fill từ kho': đề xuất fill hàng từ kho về cửa hàng\n"
                "• Sheet 'Luân chuyển cửa hàng': đề xuất chuyển giữa các cửa hàng"
            )
        except Exception as e:
            messagebox.showerror("Lỗi xuất file", str(e))


# ─── Entry ────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    app = App()
    app.mainloop()
