import os, re, sys, threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl
from openpyxl import Workbook

CODE_CJ  = 10
CODE_KD  = 16

def make_playauto_template(path):
    wb = Workbook()
    ws = wb.active
    ws.title = '작성가이드 v10.0'
    ws.append(['묶음번호', '택배사코드', '송장번호'])
    wb.save(path)

def make_icecream_template(path):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    ws.append(['배송번호', '배송순번', '택배사', '송장번호'])
    wb.save(path)

def save_playauto(rows, out):
    tmp = out + '.tmp.xlsx'
    make_playauto_template(tmp)
    wb = openpyxl.load_workbook(tmp)
    ws = wb['작성가이드 v10.0']
    ws.delete_rows(2, ws.max_row)
    for r in rows:
        ws.append([r['묶음번호'], r['택배사'], r['운송장번호']])
    wb.save(out)
    os.remove(tmp)

def save_icecream(rows, out):
    tmp = out + '.tmp.xlsx'
    make_icecream_template(tmp)
    wb = openpyxl.load_workbook(tmp)
    ws = wb['Sheet1']
    ws.delete_rows(2, ws.max_row)
    for r in rows:
        ws.append([r['배송번호'], r['배송순번'], r['택배사'], r['송장번호']])
    wb.save(out)
    os.remove(tmp)

def build_order_map(sku_df):
    m = {}
    for _, r in sku_df.iterrows():
        for oid in str(r['쇼핑몰주문번호']).split():
            m[oid.strip()] = str(r['묶음번호'])
    return m

def extract_cj(wb_df, order_map, seen=None):
    if seen is None: seen = set()
    rows, new_seen = [], set()
    for _, r in wb_df.iterrows():
        if str(r.get('상품명','')).strip() == '분리': continue
        bundle = None
        for oid in str(r['고객주문번호']).split():
            if oid.strip() in order_map:
                bundle = order_map[oid.strip()]; break
        if bundle and bundle not in seen and bundle not in new_seen:
            new_seen.add(bundle)
            rows.append({'묶음번호': bundle, '택배사': CODE_CJ, '운송장번호': str(r['운송장번호'])})
    return rows, new_seen

def extract_kd(kd_df, sku_df):
    name_map = {}
    for _, r in sku_df.iterrows():
        if str(r['SKU상품명']).startswith('(경동)'):
            name_map[str(r['수령자명']).strip()] = str(r['묶음번호'])
    rows, seen, unmatched = [], set(), []
    for _, r in kd_df.iterrows():
        name = str(r['받는분']).strip()
        bundle = name_map.get(name)
        if bundle and bundle not in seen:
            seen.add(bundle)
            rows.append({'묶음번호': bundle, '택배사': CODE_KD, '운송장번호': str(r['운송장번호'])})
        elif not bundle:
            unmatched.append(f"{name} / {r['품목명']}")
    return rows, seen, unmatched

def extract_icecream(kd_df):
    rows = []
    for _, r in kd_df.iterrows():
        if '아이스크림몰' in str(r.get('품목명', '')):
            rows.append({'배송번호': '', '배송순번': 1, '택배사': CODE_KD, '송장번호': str(r['운송장번호'])})
    return rows

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("페이퍼팝 송장 전송 자동화")
        self.geometry("680x720")
        self.resizable(False, False)
        self.configure(bg="#F5F5F5")
        self.sku_var    = tk.StringVar()
        self.wb_var     = tk.StringVar()
        self.kd_var     = tk.StringVar()
        self.prev_vars  = []
        self.output_var = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        self._build_ui()

    def _build_ui(self):
        BG="#F5F5F5"; CARD="#FFFFFF"; BLUE="#1B4F8A"
        LBLUE="#E8F0FB"; GRAY="#6B7280"; GREEN="#16A34A"; BTN_FG="#FFFFFF"; PAD=14

        def section(parent, title):
            f = tk.Frame(parent, bg=CARD, bd=0, highlightbackground="#E5E7EB", highlightthickness=1)
            f.pack(fill="x", padx=PAD, pady=(0,10))
            hdr = tk.Frame(f, bg=BLUE); hdr.pack(fill="x")
            tk.Label(hdr, text=title, bg=BLUE, fg=BTN_FG, font=("맑은 고딕",10,"bold"), anchor="w", padx=10, pady=6).pack(side="left")
            body = tk.Frame(f, bg=CARD, padx=10, pady=8); body.pack(fill="x")
            return body

        def file_row(parent, label, var, cmd):
            row = tk.Frame(parent, bg=CARD); row.pack(fill="x", pady=3)
            tk.Label(row, text=label, bg=CARD, fg=GRAY, font=("맑은 고딕",9), width=14, anchor="w").pack(side="left")
            tk.Entry(row, textvariable=var,
