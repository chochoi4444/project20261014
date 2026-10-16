"""창으로 쓰는 화면 (exe를 더블클릭하면 이 화면이 뜬다). tkinter만 쓴다."""
from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .pipeline import STATUS_KO, date_mismatch, load_mapping, run, save_report, save_workbook
from .workbook import parse_target

STATUS_ORDER = ["auto", "review", "manual", "nomatch", "skip"]
ROW_COLORS = {"auto": "#e3f3ea", "review": "#fbf0d6", "manual": "#e2efe9", "nomatch": "#ffffff", "skip": "#eceff0"}


def _fmt(n):
    return "" if n is None else f"{round(n):,}"


def open_path(p: Path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(p))  # noqa: S606 — 사용자가 누른 버튼으로 결과 폴더만 연다
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(p)])
        else:
            subprocess.Popen(["xdg-open", str(p)])
    except Exception:
        pass


class App:
    def __init__(self, root: tk.Tk, initial_paths=None):
        self.root = root
        self.q: queue.Queue = queue.Queue()
        self.ledger_paths: list[str] = []
        self.results = []
        self.out_dir: Path | None = None
        root.title(f"총자산 자동입력 {__version__}")
        root.geometry("1100x680")
        root.minsize(820, 520)
        self._build()
        if initial_paths:
            self._apply_initial(initial_paths)
        root.after(100, self._poll)

    # ---------- 화면 구성 ----------
    def _build(self):
        r = self.root
        style = ttk.Style(r)
        try:
            style.theme_use("vista" if sys.platform.startswith("win") else "clam")
        except tk.TclError:
            pass
        style.configure("Treeview", rowheight=24)
        style.configure("Big.TButton", padding=(14, 6))

        top = ttk.Frame(r, padding=(14, 12, 14, 6))
        top.pack(fill="x")
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="① 대상 엑셀").grid(row=0, column=0, sticky="w", pady=3)
        self.target_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.target_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(top, text="파일 선택…", command=self.pick_target).grid(row=0, column=2, sticky="ew")

        ttk.Label(top, text="② 기준가격대장").grid(row=1, column=0, sticky="w", pady=3)
        self.ledger_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.ledger_var, state="readonly").grid(row=1, column=1, sticky="ew", padx=6)
        btns = ttk.Frame(top)
        btns.grid(row=1, column=2, sticky="ew")
        ttk.Button(btns, text="폴더 선택…", command=self.pick_ledger_dir).pack(side="left", fill="x", expand=True)
        ttk.Button(btns, text="파일 선택…", command=self.pick_ledger_files).pack(side="left", fill="x", expand=True, padx=(4, 0))

        opt = ttk.Frame(top)
        opt.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8, 2))
        ttk.Label(opt, text="기준일").pack(side="left")
        self.date_var = tk.StringVar()
        ttk.Entry(opt, textvariable=self.date_var, width=12).pack(side="left", padx=(6, 4))
        ttk.Label(opt, text="(YYYY-MM-DD, 대상 엑셀 머리글에서 자동으로 읽음)", foreground="#66726d").pack(side="left", padx=(0, 18))
        self.review_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="'검토' 행도 입력", variable=self.review_var).pack(side="left", padx=(0, 12))
        self.overwrite_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="이미 값이 있는 칸도 덮어쓰기", variable=self.overwrite_var).pack(side="left")

        mp = ttk.Frame(top)
        mp.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(4, 2))
        mp.columnconfigure(1, weight=1)
        ttk.Label(mp, text="펀드명 매핑 (선택)").grid(row=0, column=0, sticky="w")
        self.map_var = tk.StringVar()
        ttk.Entry(mp, textvariable=self.map_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(mp, text="JSON 선택…", command=self.pick_mapping).grid(row=0, column=2)

        act = ttk.Frame(r, padding=(14, 4, 14, 6))
        act.pack(fill="x")
        self.run_btn = ttk.Button(act, text="실행 — 총자산 입력한 엑셀 저장", style="Big.TButton", command=self.start)
        self.run_btn.pack(side="left")
        self.open_btn = ttk.Button(act, text="결과 폴더 열기", command=self.open_out, state="disabled")
        self.open_btn.pack(side="left", padx=8)
        self.progress = ttk.Progressbar(act, mode="determinate", length=180)
        self.progress.pack(side="left", padx=8)
        self.status_var = tk.StringVar(value="대상 엑셀과 기준가격대장 폴더를 고른 뒤 실행을 누르세요. 원본 파일은 바뀌지 않습니다.")
        ttk.Label(act, textvariable=self.status_var).pack(side="left", padx=6)

        filt = ttk.Frame(r, padding=(14, 0, 14, 4))
        filt.pack(fill="x")
        ttk.Label(filt, text="보기").pack(side="left")
        self.filter_var = tk.StringVar(value="전체")
        cb = ttk.Combobox(filt, textvariable=self.filter_var, state="readonly", width=10,
                          values=["전체"] + [STATUS_KO[k] for k in STATUS_ORDER])
        cb.pack(side="left", padx=6)
        cb.bind("<<ComboboxSelected>>", lambda e: self._fill_table())
        self.summary_var = tk.StringVar()
        ttk.Label(filt, textvariable=self.summary_var).pack(side="left", padx=12)

        tbl = ttk.Frame(r, padding=(14, 0, 14, 12))
        tbl.pack(fill="both", expand=True)
        cols = [("row", "행", 50, "e"), ("inc", "입력", 44, "center"), ("status", "상태", 60, "center"),
                ("score", "유사도", 60, "e"), ("value", "총자산(원)", 140, "e"), ("check", "검증", 150, "w"),
                ("name", "종목명 (대상 엑셀)", 300, "w"), ("ledger", "기준가격대장", 260, "w"), ("notes", "메모", 260, "w")]
        self.tree = ttk.Treeview(tbl, columns=[c[0] for c in cols], show="headings")
        for key, title, width, anchor in cols:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor=anchor, stretch=key in ("name", "ledger", "notes"))
        for k, color in ROW_COLORS.items():
            self.tree.tag_configure(k, background=color)
        ys = ttk.Scrollbar(tbl, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(tbl, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        tbl.rowconfigure(0, weight=1)
        tbl.columnconfigure(0, weight=1)

    # ---------- 입력 ----------
    def _apply_initial(self, paths):
        paths = [p for p in paths if Path(p).exists()]
        if not paths:
            return
        first = Path(paths[0])
        if first.is_file() and first.suffix.lower() in (".xlsx", ".xlsm"):
            self._set_target(str(first))
            paths = paths[1:]
        if paths:
            self._set_ledgers(paths)

    def _set_target(self, path: str):
        self.target_var.set(path)
        try:
            t = parse_target(Path(path))
            if t.base_date:
                self.date_var.set(t.base_date)
            self.status_var.set(f"대상: 시트 '{t.sheet_name}' · {t.asset_col_letter}열 '{t.asset_header}' · 종목 {len(t.rows)}행")
        except Exception as e:
            messagebox.showerror("대상 엑셀", str(e))

    def _set_ledgers(self, paths):
        self.ledger_paths = list(paths)
        if len(paths) == 1:
            self.ledger_var.set(paths[0])
        else:
            self.ledger_var.set(f"{len(paths)}개 선택: " + ", ".join(Path(p).name for p in paths[:3]) + (" …" if len(paths) > 3 else ""))

    def pick_target(self):
        p = filedialog.askopenfilename(title="총자산을 채울 대상 엑셀", filetypes=[("Excel", "*.xlsx *.xlsm"), ("모든 파일", "*.*")])
        if p:
            self._set_target(p)

    def pick_ledger_dir(self):
        p = filedialog.askdirectory(title="기준가격대장이 들어 있는 폴더")
        if p:
            self._set_ledgers([p])

    def pick_ledger_files(self):
        ps = filedialog.askopenfilenames(title="기준가격대장 파일 (여러 개 선택 가능)",
                                         filetypes=[("Excel", "*.xlsx *.xlsm *.xls *.csv"), ("모든 파일", "*.*")])
        if ps:
            self._set_ledgers(list(ps))

    def pick_mapping(self):
        p = filedialog.askopenfilename(title="펀드명 매핑 JSON", filetypes=[("JSON", "*.json"), ("모든 파일", "*.*")])
        if p:
            self.map_var.set(p)

    # ---------- 실행 ----------
    def start(self):
        target = self.target_var.get().strip()
        if not target or not Path(target).is_file():
            messagebox.showwarning("총자산 자동입력", "대상 엑셀을 먼저 고르세요.")
            return
        if not self.ledger_paths:
            messagebox.showwarning("총자산 자동입력", "기준가격대장 폴더나 파일을 고르세요.")
            return
        base = self.date_var.get().strip() or None
        if base and not (len(base) == 10 and base[4] == "-" and base[7] == "-"):
            messagebox.showwarning("총자산 자동입력", "기준일은 2026-09-30 형식으로 적어 주세요.")
            return
        try:
            mapping = load_mapping(self.map_var.get().strip() or None)
        except Exception as e:
            messagebox.showerror("펀드명 매핑", f"매핑 파일을 읽지 못했습니다.\n{e}")
            return
        self.run_btn.configure(state="disabled")
        self.open_btn.configure(state="disabled")
        self.progress.configure(value=0)
        self.status_var.set("장부 읽는 중…")
        args = (Path(target), list(self.ledger_paths), base, self.overwrite_var.get(), self.review_var.get(), mapping)
        threading.Thread(target=self._work, args=args, daemon=True).start()

    def _work(self, target, ledgers, base, overwrite, include_review, mapping):
        try:
            def progress(i, n, f):
                self.q.put(("progress", i, n))
            t, ls, results, base_used = run(target, ledgers, base, overwrite, include_review, mapping, progress)
            stem = t.path.with_suffix("")
            report = Path(f"{stem}_매칭결과.csv")
            save_report(results, report)
            inc = sum(1 for r in results if r.include)
            out = None
            if inc:
                out = Path(f"{stem}_총자산입력.xlsx")
                save_workbook(t, results, out)
            self.q.put(("done", t, ls, results, base_used, out, report))
        except PermissionError as e:
            self.q.put(("error", f"파일을 저장하지 못했습니다. 같은 이름의 결과 파일이 엑셀에서 열려 있으면 닫고 다시 실행하세요.\n\n{e}"))
        except Exception as e:
            self.q.put(("error", f"{e}\n\n{traceback.format_exc(limit=3)}"))

    def _poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                if msg[0] == "progress":
                    _, i, n = msg
                    self.progress.configure(maximum=max(n, 1), value=i)
                    self.status_var.set(f"장부 읽는 중… {i}/{n}")
                elif msg[0] == "done":
                    self._done(*msg[1:])
                elif msg[0] == "error":
                    self.run_btn.configure(state="normal")
                    self.status_var.set("오류가 났습니다.")
                    messagebox.showerror("총자산 자동입력", msg[1])
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    def _done(self, t, ledgers, results, base, out, report):
        self.results = results
        self.run_btn.configure(state="normal")
        self.out_dir = report.parent
        self.open_btn.configure(state="normal")
        self._fill_table()
        cnt = {k: sum(1 for r in results if r.status == k) for k in STATUS_ORDER}
        inc = sum(1 for r in results if r.include)
        bad = [l for l in ledgers if l.error]
        self.summary_var.set(" · ".join(f"{STATUS_KO[k]} {cnt[k]}" for k in STATUS_ORDER) + f"   |   입력 {inc}행   |   기준일 {base or '-'}")
        lines = [f"장부 {len(ledgers)}개를 읽었습니다." + (f" (읽기 실패 {len(bad)}개)" if bad else "")]
        lines += [f"  · {l.file_name}: {l.error}" for l in bad[:5]]
        pre = [l for l in ledgers if l.superseded]
        if pre:
            lines.append(f"결산후 파일이 있어 결산전 {len(pre)}개는 쓰지 않았습니다.")
        if out:
            lines.append(f"\n총자산 {inc}칸을 채운 파일:\n{out}")
            self.status_var.set(f"저장했습니다: {out.name}")
        else:
            lines.append("\n입력할 행이 없어 엑셀은 저장하지 않았습니다.\n'검토' 행을 확인한 뒤 '검토 행도 입력'을 켜고 다시 실행하세요.")
            self.status_var.set("입력할 행이 없습니다.")
        lines.append(f"\n매칭 근거:\n{report}")
        common, n_mis = date_mismatch(results, base)
        if common:
            # 기준일이 달라 '검토'로 빠진 경우: 바로 고쳐 다시 실행할 수 있게 묻는다
            lines.append(f"\n※ 기준일은 {base}인데 매칭된 장부는 대부분 {common}자입니다.\n"
                         f"   그래서 {n_mis}행이 '검토'로 분류됐습니다.\n\n기준일을 {common}로 바꿔 다시 실행할까요?")
            if messagebox.askyesno("총자산 자동입력", "\n".join(lines)):
                self.date_var.set(common)
                self.start()
            return
        messagebox.showinfo("총자산 자동입력", "\n".join(lines))

    def _fill_table(self):
        self.tree.delete(*self.tree.get_children())
        want = self.filter_var.get()
        # 자동 → 검토 → 수동 → 미매칭 → 제외 순으로, 같은 상태 안에서는 행 번호 순
        for r in sorted(self.results, key=lambda r: (STATUS_ORDER.index(r.status), r.excel_row)):
            if want != "전체" and STATUS_KO[r.status] != want:
                continue
            L = r.ledger
            self.tree.insert("", "end", tags=(r.status,), values=(
                r.excel_row, "●" if r.include else "", STATUS_KO[r.status], f"{r.score:.2f}" if r.score else "",
                _fmt(L.value) if L else "", L.check[1] if L else "", r.name,
                (L.name or L.file_title) if L else "", " · ".join(r.notes)))

    def open_out(self):
        if self.out_dir:
            open_path(self.out_dir)


def main(initial_paths=None, selftest: bool = False) -> int:
    root = tk.Tk()
    app = App(root, initial_paths)
    if selftest:
        # 빌드 점검용: 창을 만들고 그리기까지 한 뒤 바로 닫는다
        root.update_idletasks()
        root.update()
        assert app.tree.winfo_exists()
        root.destroy()
        print("GUI selftest ok")
        return 0
    root.mainloop()
    return 0
