import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from decimal import Decimal, InvalidOperation
import pandas as pd

# Excel columns, zero-based
ERP_AMOUNT = 42       # AQ: 매출마감대상금액
ERP_TAX_TYPE = 27     # AB: 면과세
ERP_MEMBER_ID = 37    # AL: 회원ID
ERP_ORDER_TYPE = 58   # BG: 수주유형코드

MASTER_MEMBER_ID = 6  # G: 회원등록번호
MASTER_TAX = 16       # Q: 과세
MASTER_EXEMPT = 17    # R: 면세
MASTER_INVOICE = 18   # S: 인보이스

KINDS = ("과세", "면세", "인보이스")
EPSILON = Decimal("0.000001")


def cell_text(value):
    if pd.isna(value):
        return ""
    # Avoid converting an Excel numeric ID such as 12345.0 to a string with .0
    if isinstance(value, float) and value.is_integer():
        return str(int(value)).strip()
    return str(value).strip()


def member_key(value, strip_zz=False):
    key = cell_text(value)
    if strip_zz and key.startswith("ZZ_"):
        key = key[3:]
    return key.strip()


def amount(value, where=""):
    if pd.isna(value) or value == "":
        return Decimal("0")
    raw = str(value).replace(",", "").replace("원", "").strip()
    if not raw:
        return Decimal("0")
    try:
        return Decimal(raw)
    except InvalidOperation:
        raise ValueError(f"{where} 금액을 숫자로 읽을 수 없습니다: {value!r}")


def optional_numeric_amount(value):
    """마스터 금액 셀 전용: 숫자는 Decimal, 공란은 0, 텍스트는 None으로 반환."""
    if pd.isna(value) or value == "":
        return Decimal("0")
    raw = str(value).replace(",", "").replace("원", "").strip()
    if not raw:
        return Decimal("0")
    try:
        return Decimal(raw)
    except InvalidOperation:
        return None


def code_text(value):
    s = cell_text(value)
    return s[:-2] if s.endswith(".0") else s


def empty_totals():
    return {k: Decimal("0") for k in KINDS}


def read_excel(path):
    suffix = Path(path).suffix.lower()
    engine = "xlrd" if suffix == ".xls" else "openpyxl"
    return pd.read_excel(path, header=None, dtype=object, engine=engine)


def parse_erp(path):
    df = read_excel(path)
    required = max(ERP_AMOUNT, ERP_TAX_TYPE, ERP_MEMBER_ID, ERP_ORDER_TYPE)
    if df.shape[1] <= required:
        raise ValueError(f"ERP 파일에 AL열 및 BG열까지 있어야 합니다. 현재 열 수: {df.shape[1]}")

    totals = {}
    skipped_blank_id = 0
    skipped_other_code = 0
    skipped_tax_label = 0

    for row_no, row in df.iterrows():
        key = member_key(row.iloc[ERP_MEMBER_ID], strip_zz=True)
        if not key:
            skipped_blank_id += 1
            continue

        code = code_text(row.iloc[ERP_ORDER_TYPE])
        if code not in ("801", "811", "851", "861"):
            skipped_other_code += 1
            continue

        try:
            value = abs(amount(row.iloc[ERP_AMOUNT], f"ERP {row_no + 1}행 AQ열"))
        except ValueError:
            raise

        if code in ("801", "811"):
            label = cell_text(row.iloc[ERP_TAX_TYPE])
            if label == "과세":
                kind = "과세"
            elif label == "면세":
                kind = "면세"
            else:
                skipped_tax_label += 1
                continue
        else:
            kind = "인보이스"

        if code in ("811", "861"):
            value = -value

        totals.setdefault(key, empty_totals())
        totals[key][kind] += value

    return totals, {
        "blank_id": skipped_blank_id,
        "other_code": skipped_other_code,
        "tax_label": skipped_tax_label,
        "rows": len(df)
    }


def parse_master(path):
    df = read_excel(path)
    required = max(MASTER_MEMBER_ID, MASTER_TAX, MASTER_EXEMPT, MASTER_INVOICE)
    if df.shape[1] <= required:
        raise ValueError(f"마스터 파일에 S열까지 있어야 합니다. 현재 열 수: {df.shape[1]}")

    totals = {}
    skipped_blank_id = 0
    for row_no, row in df.iterrows():
        key = member_key(row.iloc[MASTER_MEMBER_ID])
        # G열의 제목 셀("회원등록번호")은 실제 회원 ID가 아니므로 건너뜁니다.
        if key == "회원등록번호":
            continue
        if not key:
            skipped_blank_id += 1
            continue
        raw_values = {
            "과세": optional_numeric_amount(row.iloc[MASTER_TAX]),
            "면세": optional_numeric_amount(row.iloc[MASTER_EXEMPT]),
            "인보이스": optional_numeric_amount(row.iloc[MASTER_INVOICE]),
        }

        # Q/R/S가 모두 텍스트인 제목·헤더 행은 합산 대상에서 제외합니다.
        # 한 칸이라도 숫자이면 나머지 텍스트/공란 셀은 0으로 간주합니다.
        if all(value is None for value in raw_values.values()):
            continue

        values = {
            kind: (Decimal("0") if value is None else value)
            for kind, value in raw_values.items()
        }
        totals.setdefault(key, empty_totals())
        for kind in KINDS:
            totals[key][kind] += values[kind]

    return totals, {"blank_id": skipped_blank_id, "rows": len(df)}


def fmt(value):
    return f"{value:,.2f}".rstrip("0").rstrip(".")


class MemberAmountReconciler:
    def __init__(self, root):
        self.root = root
        self.root.title("회원별 ERP / 마스터 금액 대조")
        self.root.geometry("1320x760")
        self.root.minsize(1050, 620)
        self.erp_path = None
        self.master_path = None
        self.erp_totals = None
        self.master_totals = None
        self.results = []
        self._ui()

    def _ui(self):
        header = ttk.Frame(self.root, padding=14)
        header.pack(fill="x")
        ttk.Label(header, text="회원별 ERP / 마스터 금액 대조", font=("맑은 고딕", 17, "bold")).pack(side="left")

        uploads = ttk.Frame(self.root, padding=(14, 2, 14, 10))
        uploads.pack(fill="x")
        erp_box = ttk.LabelFrame(uploads, text="ERP 엑셀 업로드", padding=10)
        erp_box.pack(side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Button(erp_box, text="ERP 파일 선택", command=self.select_erp).pack(anchor="w")
        self.erp_label = ttk.Label(erp_box, text="선택된 파일 없음", foreground="#555")
        self.erp_label.pack(anchor="w", pady=(6, 0))

        master_box = ttk.LabelFrame(uploads, text="마스터 금액 업로드", padding=10)
        master_box.pack(side="left", fill="x", expand=True, padx=(6, 0))
        ttk.Button(master_box, text="마스터 파일 선택", command=self.select_master).pack(anchor="w")
        self.master_label = ttk.Label(master_box, text="선택된 파일 없음", foreground="#555")
        self.master_label.pack(anchor="w", pady=(6, 0))

        rules = (
            "ERP: AL 회원ID(앞의 ZZ_ 제거) / AQ 금액 / AB 과세·면세 / BG 코드 "
            "   |   마스터: G 회원등록번호 / Q 과세 / R 면세 / S 인보이스\n"
            "주문 801·851은 양수, 취소 811·861은 음수로 집계합니다. "
            "동일 회원 ID가 여러 행이면 회원별로 합산합니다."
        )
        ttk.Label(self.root, text=rules, padding=(14, 0, 14, 10), foreground="#555").pack(anchor="w")

        controls = ttk.Frame(self.root, padding=(14, 0, 14, 8))
        controls.pack(fill="x")
        ttk.Button(controls, text="대조 시작", command=self.compare).pack(side="left")
        self.export_button = ttk.Button(controls, text="불일치 회원만 엑셀 다운로드", command=self.export_mismatches, state="disabled")
        self.export_button.pack(side="right")

        cols = ("member", "erp_tax", "master_tax", "diff_tax",
                "erp_exempt", "master_exempt", "diff_exempt",
                "erp_invoice", "master_invoice", "diff_invoice", "status")
        self.tree = ttk.Treeview(self.root, columns=cols, show="headings", height=20)
        names = {
            "member":"회원 ID", "erp_tax":"과세(ERP)", "master_tax":"과세(마스터)", "diff_tax":"과세 차액",
            "erp_exempt":"면세(ERP)", "master_exempt":"면세(마스터)", "diff_exempt":"면세 차액",
            "erp_invoice":"인보이스(ERP)", "master_invoice":"인보이스(마스터)",
            "diff_invoice":"인보이스 차액", "status":"판정"
        }
        for c in cols:
            self.tree.heading(c, text=names[c])
            self.tree.column(c, width=165 if c == "member" else (105 if c == "status" else 112),
                             anchor="w" if c == "member" else "center" if c == "status" else "e")
        sy = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        sx = ttk.Scrollbar(self.root, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.pack(fill="both", expand=True, padx=(14, 0))
        sy.place(relx=0.99, rely=0.36, relheight=0.53, anchor="ne")
        sx.pack(fill="x", padx=14)
        self.summary = ttk.Label(self.root, text="ERP 파일과 마스터 파일을 각각 업로드한 뒤 대조를 시작하세요.", padding=14)
        self.summary.pack(anchor="w")

    def select_erp(self):
        path = filedialog.askopenfilename(title="ERP 엑셀 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xlsm *.xls"), ("모든 파일", "*.*")])
        if path:
            self.erp_path = path
            self.erp_label.config(text=Path(path).name)
            self._reset_results()

    def select_master(self):
        path = filedialog.askopenfilename(title="마스터 금액 엑셀 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xlsm *.xls"), ("모든 파일", "*.*")])
        if path:
            self.master_path = path
            self.master_label.config(text=Path(path).name)
            self._reset_results()

    def _reset_results(self):
        self.erp_totals = self.master_totals = None
        self.results = []
        self.tree.delete(*self.tree.get_children())
        self.export_button.config(state="disabled")
        self.summary.config(text="파일이 변경되었습니다. 두 파일을 모두 선택한 뒤 다시 대조하세요.")

    def compare(self):
        if not self.erp_path or not self.master_path:
            messagebox.showwarning("파일 누락", "ERP 파일과 마스터 파일을 모두 업로드하세요.")
            return
        try:
            erp, erp_info = parse_erp(self.erp_path)
            master, master_info = parse_master(self.master_path)
            self.erp_totals, self.master_totals = erp, master
            self.results = []
            self.tree.delete(*self.tree.get_children())

            # 마스터 회원등록번호를 기준으로만 대조합니다. ERP에만 존재하는 회원은 제외합니다.
            keys = sorted(master.keys())
            for key in keys:
                ev = erp.get(key, empty_totals())
                mv = master.get(key, empty_totals())
                diffs = {k: ev[k] - mv[k] for k in KINDS}
                missing_side = "마스터에만 존재" if key not in erp else ""
                mismatch = bool(missing_side) or any(abs(diffs[k]) > EPSILON for k in KINDS)
                if mismatch:
                    status = missing_side or "금액 불일치"
                    result = {
                        "회원 ID": key,
                        "ERP 과세": ev["과세"], "마스터 과세": mv["과세"], "과세 차액(ERP-마스터)": diffs["과세"],
                        "ERP 면세": ev["면세"], "마스터 면세": mv["면세"], "면세 차액(ERP-마스터)": diffs["면세"],
                        "ERP 인보이스": ev["인보이스"], "마스터 인보이스": mv["인보이스"],
                        "인보이스 차액(ERP-마스터)": diffs["인보이스"],
                        "판정": status
                    }
                    self.results.append(result)
                vals = (key, fmt(ev["과세"]), fmt(mv["과세"]), fmt(diffs["과세"]),
                        fmt(ev["면세"]), fmt(mv["면세"]), fmt(diffs["면세"]),
                        fmt(ev["인보이스"]), fmt(mv["인보이스"]), fmt(diffs["인보이스"]),
                        status if mismatch else "일치")
                self.tree.insert("", "end", values=vals)

            matched = len(keys) - len(self.results)
            erp_only_count = len(set(erp) - set(master))
            self.export_button.config(state="normal" if self.results else "disabled")
            self.summary.config(
                text=f"대조 완료: 전체 회원 {len(keys):,}명 | 일치 {matched:,}명 | 불일치 {len(self.results):,}명 "
                     f"| ERP에만 존재하여 제외된 회원 {erp_only_count:,}명"
                     f" | ERP 빈 회원ID {erp_info['blank_id']}행 제외 | 마스터 빈 회원등록번호 {master_info['blank_id']}행 제외"
            )
            messagebox.showinfo("대조 완료",
                f"전체 회원: {len(keys):,}명\n불일치 회원: {len(self.results):,}명\n\n"
                f"불일치 회원만 별도 엑셀로 저장할 수 있습니다.")
        except Exception as exc:
            messagebox.showerror("대조 오류", str(exc))

    def export_mismatches(self):
        if not self.results:
            messagebox.showinfo("내보낼 데이터 없음", "불일치 회원이 없습니다.")
            return
        path = filedialog.asksaveasfilename(title="불일치 회원 저장",
            defaultextension=".xlsx", initialfile="회원별_금액불일치.xlsx",
            filetypes=[("Excel 파일", "*.xlsx")])
        if not path:
            return
        try:
            df = pd.DataFrame(self.results)
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="불일치 회원")
                ws = writer.sheets["불일치 회원"]
                ws.freeze_panes = "B2"
                ws.auto_filter.ref = ws.dimensions
                for column in ws.columns:
                    letter = column[0].column_letter
                    max_len = max(len(str(cell.value or "")) for cell in column)
                    ws.column_dimensions[letter].width = min(max(max_len + 3, 14), 30)
            messagebox.showinfo("저장 완료", f"불일치 회원 {len(self.results):,}명을 저장했습니다.\n{path}")
        except Exception as exc:
            messagebox.showerror("저장 오류", str(exc))


if __name__ == "__main__":
    root = tk.Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    app = MemberAmountReconciler(root)
    root.mainloop()
