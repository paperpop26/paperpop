import os, re, glob, sys, shutil, threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl

# ── 경로 설정 ──────────────────────────────────────────────────
def resource_path(rel):
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)

TMPL_CJ  = resource_path("template/플레이오토_대한통운_양식.xlsx")
TMPL_KD  = resource_path("template/플레이오토_경동_양식.xlsx")
TMPL_ICE = resource_path("template/아이스크림몰_송장_업로드.xlsx")

CODE_CJ  = 10
CODE_KD  = 16

# ── 핵심 로직 ──────────────────────────────────────────────────
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
            rows.append({'묶음번호': bundle, '택배사': CODE_CJ,
                         '운송장번호': str(r['운송장번호'])})
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
            rows.append({'묶음번호': bundle, '택배사': CODE_KD,
                         '운송장번호': str(r['운송장번호'])})
        elif not bundle:
            unmatched.append(f"{name} / {r['품목명']}")
    return rows, seen, unmatched

def extract_icecream(kd_df):
    rows = []
    for _, r in kd_df.iterrows():
        if '아이스크림몰' in str(r.get('품목명', '')):
            rows.append({'배송번호': '', '배송순번': 1,
                         '택배사': CODE_KD, '송장번호': str(r['운송장번호'])})
    return rows

def save_playauto(rows, tmpl, out):
    wb = openpyxl.load_workbook(tmpl)
    ws = wb['작성가이드 v10.0']
    ws.delete_rows(2, ws.max_row)
    for r in rows:
        ws.append([r['묶음번호'], r['택배사'], r['운송장번호']])
    wb.save(out)

def save_icecream(rows, tmpl, out):
    wb = openpyxl.load_workbook(tmpl)
    ws = wb['Sheet1']
    ws.delete_rows(2, ws.max_row)
    for r in rows:
        ws.append([r['배송번호'], r['배송순번'], r['택배사'], r['송장번호']])
    wb.save(out)

# ── GUI ───────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("페이퍼팝 송장 전송 자동화")
        self.geometry("680x720")
        self.resizable(False, False)
        self.configure(bg="#F5F5F5")

        # 파일 변수
        self.sku_var     = tk.StringVar()
        self.wb_var      = tk.StringVar()
        self.kd_var      = tk.StringVar()
        self.prev_vars   = []   # 미발송 파일 목록
        self.output_var  = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))

        self._build_ui()

    def _build_ui(self):
        BG     = "#F5F5F5"
        CARD   = "#FFFFFF"
        BLUE   = "#1B4F8A"
        LBLUE  = "#E8F0FB"
        GRAY   = "#6B7280"
        GREEN  = "#16A34A"
        BTN_FG = "#FFFFFF"
        PAD    = 14

        def section(parent, title):
            f = tk.Frame(parent, bg=CARD, bd=0,
                         highlightbackground="#E5E7EB", highlightthickness=1)
            f.pack(fill="x", padx=PAD, pady=(0, 10))
            hdr = tk.Frame(f, bg=BLUE)
            hdr.pack(fill="x")
            tk.Label(hdr, text=title, bg=BLUE, fg=BTN_FG,
                     font=("맑은 고딕", 10, "bold"), anchor="w",
                     padx=10, pady=6).pack(side="left")
            body = tk.Frame(f, bg=CARD, padx=10, pady=8)
            body.pack(fill="x")
            return body

        def file_row(parent, label, var, cmd):
            row = tk.Frame(parent, bg=CARD)
            row.pack(fill="x", pady=3)
            tk.Label(row, text=label, bg=CARD, fg=GRAY,
                     font=("맑은 고딕", 9), width=14, anchor="w").pack(side="left")
            tk.Entry(row, textvariable=var, font=("맑은 고딕", 9),
                     state="readonly", readonlybackground=LBLUE,
                     relief="flat", bd=1).pack(side="left", fill="x", expand=True, padx=(0,6))
            tk.Button(row, text="찾기", command=cmd, bg=BLUE, fg=BTN_FG,
                      font=("맑은 고딕", 9), relief="flat", padx=8, cursor="hand2").pack(side="right")

        # ── 타이틀 ──
        hdr_frame = tk.Frame(self, bg=BLUE, height=54)
        hdr_frame.pack(fill="x")
        hdr_frame.pack_propagate(False)
        tk.Label(hdr_frame, text="🐸  페이퍼팝 송장 전송 자동화",
                 bg=BLUE, fg=BTN_FG,
                 font=("맑은 고딕", 13, "bold")).pack(side="left", padx=16)
        tk.Label(hdr_frame,
                 text=datetime.today().strftime("%Y-%m-%d"),
                 bg=BLUE, fg="#A5C8F0",
                 font=("맑은 고딕", 10)).pack(side="right", padx=16)

        tk.Frame(self, bg=BG, height=10).pack()

        # ── 오늘 파일 ──
        s1 = section(self, "  📂  오늘 파일")
        file_row(s1, "SKU 매칭명",   self.sku_var, self._pick_sku)
        file_row(s1, "운송장출력데이터", self.wb_var,  self._pick_wb)
        file_row(s1, "경동 발송자료", self.kd_var,  self._pick_kd)

        # ── 미발송 파일 ──
        s2 = section(self, "  📂  미발송 운송장 파일 (선택)")
        self.prev_listbox_frame = tk.Frame(s2, bg=CARD)
        self.prev_listbox_frame.pack(fill="x")
        self.prev_listbox = tk.Listbox(self.prev_listbox_frame,
                                       font=("맑은 고딕", 9), height=3,
                                       bg=LBLUE, relief="flat", bd=0,
                                       selectbackground=BLUE, selectforeground="white")
        self.prev_listbox.pack(side="left", fill="x", expand=True)
        sb = ttk.Scrollbar(self.prev_listbox_frame, orient="vertical",
                           command=self.prev_listbox.yview)
        sb.pack(side="right", fill="y")
        self.prev_listbox.config(yscrollcommand=sb.set)
        btn_row = tk.Frame(s2, bg=CARD)
        btn_row.pack(fill="x", pady=(4,0))
        tk.Button(btn_row, text="+ 추가", command=self._add_prev,
                  bg=BLUE, fg=BTN_FG, font=("맑은 고딕", 9),
                  relief="flat", padx=10, cursor="hand2").pack(side="left", padx=(0,6))
        tk.Button(btn_row, text="삭제", command=self._del_prev,
                  bg="#9CA3AF", fg=BTN_FG, font=("맑은 고딕", 9),
                  relief="flat", padx=10, cursor="hand2").pack(side="left")

        # ── 저장 경로 ──
        s3 = section(self, "  💾  저장 경로")
        file_row(s3, "출력 폴더", self.output_var, self._pick_output)

        # ── 실행 버튼 ──
        run_frame = tk.Frame(self, bg=BG)
        run_frame.pack(fill="x", padx=PAD, pady=4)
        self.run_btn = tk.Button(run_frame, text="▶  송장 파일 만들기",
                                 command=self._run,
                                 bg=GREEN, fg=BTN_FG,
                                 font=("맑은 고딕", 12, "bold"),
                                 relief="flat", pady=10, cursor="hand2")
        self.run_btn.pack(fill="x")

        # ── 로그 ──
        s4 = section(self, "  📋  실행 로그")
        self.log_text = tk.Text(s4, height=10, font=("맑은 고딕", 9),
                                bg="#F9FAFB", relief="flat", bd=0,
                                state="disabled", wrap="word")
        self.log_text.pack(fill="both")
        self.log_text.tag_config("ok",   foreground=GREEN)
        self.log_text.tag_config("warn", foreground="#D97706")
        self.log_text.tag_config("err",  foreground="#DC2626")
        self.log_text.tag_config("info", foreground=BLUE)

    # ── 파일 선택 ──────────────────────────────────────────────
    def _pick(self, var, title):
        p = filedialog.askopenfilename(title=title,
            filetypes=[("Excel 파일", "*.xlsx")])
        if p: var.set(p)

    def _pick_sku(self):    self._pick(self.sku_var,    "SKU 매칭명 파일 선택")
    def _pick_wb(self):     self._pick(self.wb_var,     "운송장출력데이터 파일 선택")
    def _pick_kd(self):     self._pick(self.kd_var,     "경동 발송자료 파일 선택")
    def _pick_output(self):
        p = filedialog.askdirectory(title="저장 폴더 선택")
        if p: self.output_var.set(p)

    def _add_prev(self):
        paths = filedialog.askopenfilenames(title="미발송 운송장 파일 선택 (여러 개 가능)",
                                           filetypes=[("Excel 파일", "*.xlsx")])
        for p in paths:
            if p not in self.prev_vars:
                self.prev_vars.append(p)
                self.prev_listbox.insert("end", os.path.basename(p))

    def _del_prev(self):
        sel = self.prev_listbox.curselection()
        if sel:
            idx = sel[0]
            self.prev_vars.pop(idx)
            self.prev_listbox.delete(idx)

    # ── 로그 출력 ──────────────────────────────────────────────
    def _log(self, msg, tag=""):
        self.log_text.config(state="normal")
        self.log_text.insert("end", msg + "\n", tag)
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        self.update_idletasks()

    def _clear_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    # ── 실행 ───────────────────────────────────────────────────
    def _run(self):
        if not self.sku_var.get() or not self.wb_var.get() or not self.kd_var.get():
            messagebox.showwarning("파일 미선택", "오늘 파일 3개를 모두 선택해 주세요.")
            return
        self.run_btn.config(state="disabled", text="처리 중...")
        threading.Thread(target=self._process, daemon=True).start()

    def _process(self):
        try:
            self._clear_log()
            today = datetime.today().strftime('%Y%m%d')
            self._log(f"[{today}] 처리 시작", "info")

            sku_df = pd.read_excel(self.sku_var.get(), sheet_name=0)
            wb_df  = pd.read_excel(self.wb_var.get())
            kd_df  = pd.read_excel(self.kd_var.get())
            self._log("✔ 파일 읽기 완료", "ok")

            order_map = build_order_map(sku_df)

            # 대한통운 오늘
            cj_today, seen_today = extract_cj(wb_df, order_map)
            all_seen = set(seen_today)

            # 미발송 누적
            prev_rows = []
            for path in self.prev_vars:
                df = pd.read_excel(path)
                rows, seen = extract_cj(df, order_map, already_seen=all_seen)
                all_seen |= seen
                prev_rows.extend(rows)
                self._log(f"  + 미발송 {os.path.basename(path)}: {len(rows)}건")

            # 경동
            kd_rows, kd_seen, unmatched = extract_kd(kd_df, sku_df)

            # 경동 우선: CJ에서 제거
            cj_final = [r for r in (cj_today + prev_rows)
                        if r['묶음번호'] not in kd_seen]

            # 아이스크림몰
            ice_rows = extract_icecream(kd_df)

            self._log(f"✔ 대한통운: {len(cj_final)}건 (오늘 {len(cj_today)} + 미발송 {len(prev_rows)})", "ok")
            self._log(f"✔ 경동: {len(kd_rows)}건", "ok")

            if unmatched:
                self._log(f"⚠ 경동 미매칭 {len(unmatched)}건 (수동 추가 필요)", "warn")
                for u in unmatched:
                    self._log(f"   {u}", "warn")

            if ice_rows:
                self._log(f"✔ 아이스크림몰: {len(ice_rows)}건 (배송번호 수동 입력 필요)", "warn")

            # 저장
            out_dir = self.output_var.get()
            os.makedirs(out_dir, exist_ok=True)

            out_main = os.path.join(out_dir, f"2026_플레이오토_송신_{today}.xlsx")
            save_playauto(cj_final + kd_rows, TMPL_CJ, out_main)
            self._log(f"✔ 저장: {os.path.basename(out_main)}", "ok")

            if ice_rows:
                out_ice = os.path.join(out_dir, f"아이스크림몰_송장_{today}.xlsx")
                save_icecream(ice_rows, TMPL_ICE, out_ice)
                self._log(f"✔ 저장: {os.path.basename(out_ice)}", "ok")

            self._log("\n완료! 출력 폴더를 확인해 주세요 🐸", "ok")

            # 출력 폴더 열기
            os.startfile(out_dir) if sys.platform == "win32" else os.system(f'open "{out_dir}"')

        except Exception as e:
            self._log(f"오류 발생: {e}", "err")
        finally:
            self.run_btn.config(state="normal", text="▶  송장 파일 만들기")


if __name__ == "__main__":
    app = App()
    app.mainloop()
