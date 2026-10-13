"""Windows desktop UI, using only the Python standard library."""

from datetime import date, timedelta
import json
import os
from pathlib import Path
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, font, messagebox, simpledialog, ttk

from .generator import generate_files, load_schedule, presentation_eligible
from .service import analyze_files, apply_original, display_value, export_result, load_journal
from .workbook import Workbook
from . import __version__
from .settings import load_settings, save_settings


def asset_root():
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent)) / "attendance_app" / "assets"


class Application(tk.Tk):
    def __init__(self, settings_file=None):
        super().__init__()
        self.title(f"출석부 · 산출물 유틸리티 {__version__}")
        screen_width, screen_height = self.winfo_screenwidth(), self.winfo_screenheight()
        width, height = min(1140, screen_width - 60), min(900, screen_height - 100)
        self.geometry(f"{width}x{height}+{max(0, (screen_width - width) // 2)}+{max(0, (screen_height - height - 40) // 2)}")
        self.minsize(min(1000, width), min(740, height))
        self.configure(bg="#f3f5f9")
        self.data = None
        self.settings_file = settings_file
        self.saved_settings = load_settings(settings_file)
        self.events = {}
        self.event_tables = {}
        self.generate_summaries = {}
        self.zip_flags = {}
        self.busy = False
        self.last_folder = None
        self.messages = queue.Queue()
        self.variables = {}
        self.action_buttons = []
        self.generation_signatures = {}
        self.analysis_signature = None
        self._styles()
        self._build()
        self.after(100, self._poll)
        self.protocol("WM_DELETE_WINDOW", self.close)

    def _styles(self):
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            font.nametofont(name).configure(family="맑은 고딕", size=10)
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", font=("맑은 고딕", 10))
        style.configure("TFrame", background="#f3f5f9")
        style.configure("TLabel", background="#f3f5f9", foreground="#233149")
        style.configure("Title.TLabel", font=("맑은 고딕", 20, "bold"))
        style.configure("Hint.TLabel", foreground="#5c6d83", font=("맑은 고딕", 9))
        style.configure("TButton", font=("맑은 고딕", 10), padding=(12, 7), background="#e8edf4", foreground="#233149")
        style.map("TButton", background=[("active", "#d7e2f0")])
        style.configure("Primary.TButton", background="#255b9e", foreground="white")
        style.map("Primary.TButton", background=[("active", "#1e487d"), ("disabled", "#9aaabd")])
        style.configure("TEntry", padding=6)
        style.configure("TNotebook", background="#f3f5f9", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(18, 9), background="#e3e9f2", foreground="#4b5d76")
        style.map("TNotebook.Tab", background=[("selected", "#ffffff"), ("active", "#d7e2f0")], foreground=[("selected", "#255b9e")])
        style.configure("TLabelframe", background="#f3f5f9", bordercolor="#d5dce6")
        style.configure("TLabelframe.Label", background="#f3f5f9", foreground="#233149", font=("맑은 고딕", 10, "bold"))
        style.configure("Treeview", rowheight=29, font=("맑은 고딕", 10), background="white", fieldbackground="white")
        style.configure("Treeview.Heading", font=("맑은 고딕", 10, "bold"), padding=6)
        style.map("Treeview", background=[("selected", "#d7e7fa")], foreground=[("selected", "#16395f")])

    def var(self, name, default=""):
        if name not in self.variables:
            saved = self.saved_settings.get(name)
            if saved and name != "output" and not Path(saved).is_file():
                saved = None
            self.variables[name] = tk.StringVar(self, value=saved or default)
        return self.variables[name]

    def value(self, name):
        return self.variables[name].get().strip()

    def button(self, parent, text, command, primary=False):
        button = ttk.Button(parent, text=text, command=command, style="Primary.TButton" if primary else "TButton")
        self.action_buttons.append(button)
        return button

    def file_row(self, parent, row, label, name, extension=None, folder=False, default=""):
        ttk.Label(parent, text=label, width=17).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Entry(parent, textvariable=self.var(name, default)).grid(row=row, column=1, sticky="ew", pady=4)
        def choose():
            path = filedialog.askdirectory(parent=self) if folder else filedialog.askopenfilename(parent=self, filetypes=[(extension or "파일", "*." + extension if extension else "*.*")])
            if path:
                self.var(name).set(path)
                if name in ("output", "attendance_template", "ppt_template", "schedule"):
                    self.remember_settings()
        self.button(parent, "폴더 선택" if folder else "파일 선택", choose).grid(row=row, column=2, padx=(8, 0), pady=4)
        parent.columnconfigure(1, weight=1)

    def table(self, parent, columns, height=8):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings", height=height, selectmode="extended")
        for name, label, width in columns:
            tree.heading(name, text=label)
            tree.column(name, width=width, minwidth=60, anchor="w", stretch=True)
        vertical = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        return tree

    def _build(self):
        shell = ttk.Frame(self, padding=18)
        shell.pack(fill="both", expand=True)
        ttk.Label(shell, text="출석부 · 산출물", style="Title.TLabel").pack(anchor="w")
        ttk.Label(shell, text="파일 준비부터 출결 확인까지, 한 곳에서 처리합니다.", style="Hint.TLabel").pack(anchor="w", pady=(2, 8))
        output_row = ttk.Frame(shell)
        output_row.pack(fill="x", pady=(0, 8))
        self.file_row(output_row, 0, "공통 출력 폴더", "output", folder=True, default=str(Path.home() / "Downloads" / "출석부_산출물"))
        self.button(output_row, "출력 폴더 열기", self.open_output_folder).grid(row=0, column=3, padx=(8, 0))
        self.status = tk.StringVar(self, value="입력 파일을 선택하세요.")
        ttk.Label(shell, textvariable=self.status, style="Hint.TLabel", wraplength=min(1040, self.winfo_screenwidth() - 100)).pack(side="bottom", anchor="w", fill="x", pady=(10, 0))
        notebook = ttk.Notebook(shell)
        self.notebook = notebook
        notebook.pack(fill="both", expand=True)
        self.update_tab = ttk.Frame(notebook, padding=14)
        self.attendance_tab = ttk.Frame(notebook, padding=14)
        self.presentation_tab = ttk.Frame(notebook, padding=14)
        help_tab = ttk.Frame(notebook, padding=20)
        notebook.add(self.update_tab, text="출석부 업데이트")
        notebook.add(self.attendance_tab, text="출석부 생성")
        notebook.add(self.presentation_tab, text="산출물 생성")
        notebook.add(help_tab, text="사용 안내")
        self._update_ui()
        self._generate_ui(self.attendance_tab, "attendance")
        self._generate_ui(self.presentation_tab, "presentation")
        instructions = """출석부 업데이트
1. 출석부 Excel과 Teams 참석 보고서 CSV를 선택합니다. 중간출결이 없는 날은 CSV를 비워 둡니다.
2. 일반 수업 또는 프로젝트 운영 방식을 선택하고 확인 시간을 입력한 뒤 ‘분석하기’를 누릅니다.
3. 변경 셀, 외출 후보, 확인할 기록을 살펴보고 ‘업데이트 파일 저장’을 누릅니다.
4. ‘변경 셀 원본 반영’을 누르면 선택한 출석부를 백업한 뒤 변경값을 반영합니다.
   저장한 변경내역.json은 다음 실행 때 불러올 수 있습니다.

출석부 생성 / 산출물 생성
1. 필요한 탭을 선택합니다. 출석부와 산출물은 각각 따로 생성합니다.
2. 출석부 양식·명단과 교육 일정 CSV, 기간을 지정하고 ‘일정 불러오기’를 누릅니다.
3. 출석부는 생성 대상 교육 일정에, 산출물은 강사가 지정된 수업 일정에만 만듭니다.
   산출물에서는 프로젝트·타운홀 등 수업 외 일정과 강사 미정 일정을 제외합니다.
4. 중도퇴소자가 생기면 출석부 양식에서 명단을 수정하고 다시 선택·불러오기 하세요.
   산출물 PPT도 그 명단과 순번에 맞춰 앞 4장 + 학생별 4장으로 조정합니다.

공통 출력 폴더
• 화면 위에서 한 번 지정하면 다음 실행에도 기억합니다.
• 업데이트 파일·검토 보고서·원본 반영 결과·백업·생성 파일을 모두 이 폴더 아래에 모읍니다.
• 원본 반영은 원본의 기존 위치에서 수정하며, 공통 출력 폴더에도 반영 결과 복사본을 저장합니다.

처리 기준
• 업데이트: H~J Teams 원문 시간, K 참석 O·미참석 x. LMS·호명·특이사항·확정출결·비고는 유지합니다.
• 저장 파일과 원본 반영 모두 변경 셀을 주황색으로 강조합니다. 값만 반영된 셀은 강조색을 보완합니다.
• 업데이트 파일 저장은 기존 웹앱처럼 증빙 이미지를 제외하며, 원본 반영과 반영 결과 복사본은 이미지를 유지합니다.
• 외출 후보: 휴게시간 제외 후 10분 이상 연속 접속 공백. 수기 판정을 자동 확정하지 않습니다.
• 점심은 기본 12:50~14:00이며, 실제 시작이 이르면 수정합니다. 시간별 휴게는 직접 입력합니다.
• 프로젝트 모드는 전체 모임방의 오전·종료 접속만 확인하고, 팀 이동 시간은 외출로 계산하지 않습니다.
• 프로젝트 팀 참석 보고서가 없으므로 낮 동안의 참여 여부는 자동 판정하지 않습니다.
• 최종 퇴장 후 재입장이 없는 기록은 외출로 계산하지 않습니다.
• 이름이 중복되거나 날짜가 다른 파일은 자동 연결하지 않습니다.
• 새 출석부 생성 시 학생별 당일 입력값과 증빙 이미지·댓글을 비웁니다.
• 원본 반영 전후에는 Excel에서 해당 파일을 닫아 두세요. 백업은 공통 출력 폴더의 ‘출석부_백업’에 저장합니다.
• 일정 CSV에서 file_required=no 및 공휴일로 명시된 항목은 제외합니다. 생성 전 목록을 확인하세요.
• 모든 처리는 이 컴퓨터에서 수행합니다.
"""
        text = tk.Text(help_tab, wrap="word", relief="flat", bg="white", fg="#233149", padx=18, pady=16)
        text.insert("1.0", instructions)
        text.configure(state="disabled")
        text.pack(fill="both", expand=True)

    def _update_ui(self):
        inputs = ttk.LabelFrame(self.update_tab, text="입력 파일", padding=12)
        inputs.pack(fill="x")
        self.file_row(inputs, 0, "출석부", "workbook", "xlsx")
        self.file_row(inputs, 1, "중간출결 (선택)", "middle", "csv")
        self.button(inputs, "중간출결 안 함", lambda: self.var("middle").set("")).grid(row=1, column=3, padx=(8, 0))
        self.file_row(inputs, 2, "Teams 참석 보고서", "teams", "csv")
        ttk.Label(inputs, text="4시간 수업 등 중간출결을 하지 않는 날은 중간출결 칸을 비워 두세요. K열 유지·누락 경고 제외", style="Hint.TLabel").grid(row=3, column=0, columnspan=3, sticky="w", pady=(5, 0))
        times = ttk.LabelFrame(self.update_tab, text="운영 방식과 확인 시간", padding=12)
        times.pack(fill="x", pady=(10, 8))
        ttk.Label(times, text="운영 방식").grid(row=0, column=0, sticky="w", padx=(0, 8))
        mode = ttk.Combobox(times, textvariable=self.var("update_mode", "일반 수업"), values=("일반 수업", "프로젝트"), state="readonly", width=12)
        mode.grid(row=0, column=1, sticky="w", padx=(0, 20))
        for col, (label, name, default) in enumerate((("수업 시작", "class_start", "09:10"), ("수업 종료", "class_end", "17:50")), 1):
            ttk.Label(times, text=label).grid(row=0, column=col * 2, padx=(0, 8))
            ttk.Entry(times, textvariable=self.var(name, default), width=9).grid(row=0, column=col * 2 + 1, sticky="w", padx=(0, 20))
        ttk.Label(times, text="점심").grid(row=1, column=0, sticky="w", pady=(8, 0))
        lunch_entry = ttk.Entry(times, textvariable=self.var("lunch", "12:50-14:00"), width=16)
        lunch_entry.grid(row=1, column=1, sticky="w", pady=(8, 0), padx=(0, 20))
        ttk.Label(times, text="추가 휴게").grid(row=1, column=2, sticky="w", pady=(8, 0), padx=(0, 8))
        breaks_entry = ttk.Entry(times, textvariable=self.var("breaks"))
        breaks_entry.grid(row=1, column=3, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Label(times, text="오전 확인 종료").grid(row=2, column=0, sticky="w", pady=(8, 0))
        morning_entry = ttk.Entry(times, textvariable=self.var("morning_end", "09:20"), width=9)
        morning_entry.grid(row=2, column=1, sticky="w", pady=(8, 0), padx=(0, 20))
        ttk.Label(times, text="종료 확인 시작").grid(row=2, column=2, sticky="w", pady=(8, 0), padx=(0, 8))
        final_entry = ttk.Entry(times, textvariable=self.var("final_start", "17:40"), width=9)
        final_entry.grid(row=2, column=3, sticky="w", pady=(8, 0))
        self.mode_hint = tk.StringVar(self)
        ttk.Label(times, textvariable=self.mode_hint, style="Hint.TLabel", wraplength=1010).grid(row=3, column=0, columnspan=6, sticky="w", pady=(6, 0))
        times.columnconfigure(5, weight=1)
        self.regular_time_entries = (lunch_entry, breaks_entry)
        self.project_time_entries = (morning_entry, final_entry)
        self.var("update_mode").trace_add("write", lambda *_: self.refresh_analysis_mode())
        self.refresh_analysis_mode()
        actions = ttk.Frame(self.update_tab)
        actions.pack(fill="x", pady=(0, 8))
        self.button(actions, "분석하기", self.analyze, True).pack(side="left")
        self.button(actions, "업데이트 파일 저장", self.export).pack(side="left", padx=6)
        self.button(actions, "변경 셀 원본 반영", self.apply).pack(side="left")
        self.button(actions, "변경내역 불러오기", self.load).pack(side="right")
        self.summary = tk.StringVar(self, value="운영 방식을 선택하고 분석하면 출석 확인 결과가 표시됩니다.")
        ttk.Label(self.update_tab, textvariable=self.summary, wraplength=1010).pack(anchor="w", pady=(0, 8))
        result_tabs = ttk.Notebook(self.update_tab)
        result_tabs.pack(fill="both", expand=True)
        self.tables = {}
        for key, title, cols in (
            ("changes", "변경 셀", [("name", "이름", 105), ("cell", "셀", 70), ("before", "기존 값", 140), ("after", "변경 값", 140), ("source", "근거", 230)]),
            ("project", "프로젝트 출석 확인", [("name", "이름", 95), ("morning", "오전", 75), ("final", "종료", 75), ("middle", "중간", 70), ("first", "첫 입장", 200), ("last", "마지막 퇴장", 200)]),
            ("gaps", "외출 확인 후보", [("name", "이름", 120), ("start", "시작", 160), ("end", "종료", 160), ("duration", "연속 공백", 160)]),
            ("raw", "전체 접속 공백", [("name", "이름", 120), ("start", "퇴장", 160), ("end", "재입장", 160), ("duration", "휴게 제외 전", 160)]),
            ("findings", "확인할 기록", [("name", "이름", 100), ("kind", "구분", 210), ("detail", "내용", 580)]),
        ):
            tab = ttk.Frame(result_tabs, padding=4)
            result_tabs.add(tab, text=title)
            self.tables[key] = self.table(tab, cols, height=7)

    def _generate_ui(self, parent, kind):
        root = asset_root()
        is_ppt = kind == "presentation"
        inputs = ttk.LabelFrame(parent, text="양식과 교육 일정", padding=12)
        inputs.pack(fill="x")
        self.file_row(inputs, 0, "명단 출석부" if is_ppt else "출석부 양식·명단", "attendance_template", "xlsx", default=str(root / "출석부_양식.xlsx"))
        if is_ppt:
            self.file_row(inputs, 1, "산출물 PPT 양식", "ppt_template", "pptx", default=str(root / "산출물_양식.pptx"))
        self.file_row(inputs, 2 if is_ppt else 1, "교육 일정", "schedule", "csv", default=str(root / "교육일정.csv"))
        note = "강사가 지정된 수업만 산출물을 생성합니다. 강사 미정·프로젝트·타운홀 등은 제외합니다." if is_ppt else "중도퇴소자가 발생하면 명단을 수정한 출석부 양식을 선택하고 일정을 다시 불러오세요."
        ttk.Label(parent, text=note, style="Hint.TLabel", wraplength=1010).pack(anchor="w", pady=(8, 0))
        options = ttk.Frame(parent)
        options.pack(fill="x", pady=10)
        today = date.today()
        first = today.replace(day=1)
        following = (first.replace(day=28) + timedelta(days=4)).replace(day=1)
        for label, name, default in (("시작일", kind + "_first", first.isoformat()), ("종료일", kind + "_last", (following - timedelta(days=1)).isoformat())):
            ttk.Label(options, text=label).pack(side="left", padx=(0, 6))
            ttk.Entry(options, textvariable=self.var(name, default), width=13).pack(side="left", padx=(0, 14))
        self.button(options, "일정 불러오기", lambda: self.read_schedule(kind)).pack(side="left")
        self.button(options, "전체 선택", lambda: self.event_tables[kind].selection_set(self.event_tables[kind].get_children())).pack(side="right")
        self.generate_summaries[kind] = tk.StringVar(self, value="양식과 기간을 지정한 뒤 일정을 불러오세요.")
        ttk.Label(parent, textvariable=self.generate_summaries[kind], style="Hint.TLabel", wraplength=1010).pack(anchor="w", pady=(0, 8))
        actions = ttk.Frame(parent)
        actions.pack(side="bottom", fill="x", pady=(10, 0))
        self.zip_flags[kind] = tk.BooleanVar(self, value=True)
        ttk.Checkbutton(actions, text="ZIP도 생성", variable=self.zip_flags[kind]).pack(side="left")
        self.button(actions, "선택 일정 산출물 생성" if is_ppt else "선택 일정 출석부 생성", lambda: self.generate(kind), True).pack(side="left", padx=12)
        self.button(actions, "결과 폴더 열기", self.open_folder).pack(side="right")
        self.event_tables[kind] = self.table(parent, [("date", "날짜", 115), ("event", "수업·행사", 460), ("instructor", "강사 (더블클릭 수정)", 170), ("type", "유형", 90)], height=6)
        self.event_tables[kind].bind("<Double-1>", lambda event: self.edit_instructor(event, kind))

    def remember_settings(self):
        values = {"output": self.value("output")}
        for key in ("attendance_template", "ppt_template", "schedule"):
            if key in self.variables and self.value(key):
                path = Path(self.value(key)).resolve()
                if not path.is_relative_to(asset_root().resolve()):
                    values[key] = str(path)
        if values != self.saved_settings:
            try:
                save_settings(values, self.settings_file)
                self.saved_settings = values
            except OSError as error:
                messagebox.showwarning("설정 저장", f"현재 설정은 사용할 수 있지만 다음 실행을 위한 저장에 실패했습니다.\n{error}", parent=self)

    def output_folder(self):
        value = self.value("output")
        if not value or Path(value).is_file():
            messagebox.showinfo("출력 폴더", "화면 위에서 공통 출력 폴더를 지정하세요.", parent=self)
            return None
        self.remember_settings()
        return str(Path(value).resolve())

    def open_output_folder(self):
        folder = self.output_folder()
        if folder:
            try:
                Path(folder).mkdir(parents=True, exist_ok=True)
                os.startfile(folder)
            except OSError as error:
                messagebox.showerror("출력 폴더", str(error), parent=self)

    def run(self, message, function, on_success):
        if self.busy:
            return
        self.busy = True
        self.status.set(message)
        for button in self.action_buttons:
            button.configure(state="disabled")
        def work():
            try:
                self.messages.put((True, function(), on_success))
            except Exception as error:
                self.messages.put((False, str(error), None))
        threading.Thread(target=work, daemon=True).start()

    def _poll(self):
        try:
            success, result, callback = self.messages.get_nowait()
        except queue.Empty:
            pass
        else:
            self.busy = False
            for button in self.action_buttons:
                button.configure(state="normal")
            if success:
                callback(result)
            else:
                self.status.set("처리하지 못했습니다. 입력 파일과 안내 내용을 확인하세요.")
                messagebox.showerror("처리 확인", result, parent=self)
        self.after(100, self._poll)

    def refresh_analysis_mode(self):
        project = self.value("update_mode") == "프로젝트"
        for entry in self.regular_time_entries:
            entry.configure(state="disabled" if project else "normal")
        for entry in self.project_time_entries:
            entry.configure(state="normal" if project else "disabled")
        self.mode_hint.set("전체 모임방의 오전·종료 접속만 확인합니다. 프로젝트 팀 이동 시간은 외출로 계산하지 않습니다." if project else "실제 휴게시간을 쉼표로 구분하세요: 09:45-09:55, 10:50-11:00 · 미입력 휴게는 제외하지 않습니다.")

    def analysis_inputs(self):
        return tuple(self.value(n) for n in ("workbook", "middle", "teams", "class_start", "class_end", "lunch", "breaks", "update_mode", "morning_end", "final_start"))

    def analyze(self):
        inputs = self.analysis_inputs()
        breaks = [b.strip() for b in inputs[6].replace(";", ",").split(",") if b.strip()]
        mode = "project" if inputs[7] == "프로젝트" else "regular"
        def done(data):
            self.analysis_signature = inputs
            self.show_data(data)
            self.status.set("분석 완료. 결과를 확인한 뒤 업데이트 파일을 저장하세요.")
        self.run("입력 파일을 읽고 출결을 연결하고 있습니다…", lambda: analyze_files(*inputs[:6], breaks, mode, inputs[8], inputs[9]), done)

    def show_data(self, data):
        self.data = data
        for tree in self.tables.values():
            tree.delete(*tree.get_children())
        for c in data["proposed_changes"]:
            self.tables["changes"].insert("", "end", values=(c["name"], c["cell"], display_value(c["before"], c["type"]), display_value(c["after"], c["type"]), c["source"]))
        for check in data.get("project_attendance_checks", []):
            self.tables["project"].insert("", "end", values=(check["name"], check["morning"], check["final"], check["middle"], check["first"], check["last"]))
        for key, field in (("gaps", "class_gap_candidates"), ("raw", "disconnection_candidates")):
            for g in data[field]:
                self.tables[key].insert("", "end", values=(g["name"], g["start"][11:], g["end"][11:], g["duration"]))
        for finding in data["findings"]:
            self.tables["findings"].insert("", "end", values=(finding["name"], finding["kind"], finding["detail"]))
        for key, label in (("middle_outside_roster", "중간출결"), ("teams_outside_roster", "Teams")):
            for name in data["matching"][key]:
                self.tables["findings"].insert("", "end", values=(name, "명단 밖 인원", f"{label}에만 있어 출석부에 추가하지 않음"))
        if not data["schedule"].get("middle_enabled", True):
            self.tables["findings"].insert("", "end", values=("전체", "중간출결 해당 없음", "CSV 미사용 · K열 유지 · 누락 경고 제외"))
        if data["schedule"].get("mode", "regular") == "project":
            summary = data.get("project_summary", {})
            self.summary.set(f"{data['date']} · 프로젝트 · 수강생 {data['workbook']['roster_count']}명 · 오전 미확인 {summary.get('morning_unconfirmed', 0)}명 · 종료 미확인 {summary.get('final_unconfirmed', 0)}명 · 낮 동안 외출 분석 제외")
        else:
            self.summary.set(f"{data['date']} · 일반 수업 · 수강생 {data['workbook']['roster_count']}명 · 변경 {len(data['proposed_changes'])}셀 · 외출 후보 {len(data['class_gap_candidates'])}건 (실제 휴게시간 확인 필요)")

    def require_data(self):
        if not self.data:
            messagebox.showinfo("분석 필요", "먼저 분석하거나 저장한 변경내역을 불러오세요.", parent=self)
            return False
        if self.analysis_signature != self.analysis_inputs():
            messagebox.showinfo("재분석 필요", "입력 파일 또는 수업·휴게시간이 바뀌었습니다. 다시 분석하세요.", parent=self)
            return False
        return True

    def export(self):
        if not self.require_data():
            return
        folder = self.output_folder()
        if not folder:
            return
        def done(result):
            self.last_folder = result["folder"]
            self.status.set(f"{result['count']}개 셀 반영한 복사본 저장: {result['workbook']}")
            messagebox.showinfo("저장 완료", "업데이트 출석부, 출결검토.html, 변경내역.json을 저장했습니다.\n\n" + result["folder"], parent=self)
        self.run("업데이트 파일과 검토 결과를 저장하고 있습니다…", lambda: export_result(self.data, folder), done)

    def apply(self):
        if not self.require_data():
            return
        folder = self.output_folder()
        if not folder:
            return
        target = filedialog.askopenfilename(parent=self, title="변경 셀을 반영할 기존 출석부", initialdir=str(Path(self.data["source_workbook"]).parent), initialfile=Path(self.data["source_workbook"]).name, filetypes=[("Excel", "*.xlsx")])
        if not target:
            return
        if not messagebox.askokcancel("원본 반영", f"{Path(target).name}\n\n최대 {len(self.data['proposed_changes'])}개 셀에 값과 주황색 강조를 반영합니다.\n백업·반영 결과·검토 기록은 공통 출력 폴더에 모읍니다.\n{folder}\n\nExcel에서 원본 파일을 닫은 뒤 진행하세요.", parent=self):
            return
        def done(result):
            if result["folder"]:
                self.last_folder = result["folder"]
            self.status.set(f"원본 반영 완료: {result['count']}개 셀")
            messagebox.showinfo("반영 완료", f"{result['count']}개 셀을 반영했습니다.\n\n백업: {result['backup'] or '추가 변경 없음'}\n결과 폴더: {result['folder'] or folder}", parent=self)
        self.run("기존 값과 충돌을 확인하고 백업 후 반영하고 있습니다…", lambda: apply_original(self.data, target, folder), done)

    def load(self):
        path = filedialog.askopenfilename(parent=self, title="저장한 변경내역 선택", filetypes=[("변경내역", "*.json")])
        if not path:
            return
        try:
            data = load_journal(path)
            self.var("workbook").set(data["source_workbook"])
            self.var("middle").set(data["sources"][1]["path"])
            self.var("teams").set(data["sources"][2]["path"])
            for key in ("class_start", "class_end", "lunch"):
                self.var(key).set(data["schedule"][key])
            self.var("breaks").set(", ".join(data["schedule"]["additional_breaks"]))
            self.var("update_mode").set("프로젝트" if data["schedule"].get("mode", "regular") == "project" else "일반 수업")
            self.var("morning_end").set(data["schedule"].get("morning_end", "09:20"))
            self.var("final_start").set(data["schedule"].get("final_start", "17:40"))
            self.refresh_analysis_mode()
            self.analysis_signature = self.analysis_inputs()
            self.show_data(data)
            self.status.set("저장한 분석 결과를 불러왔습니다. 원본 반영 시 현재 셀과 다시 비교합니다.")
        except Exception as error:
            messagebox.showerror("불러오기 실패", str(error), parent=self)

    def schedule_inputs(self, kind):
        names = ("attendance_template", "ppt_template", "schedule", kind + "_first", kind + "_last")
        values = tuple(self.value(n) if n in self.variables else "" for n in names)
        if kind == "attendance":
            values = (values[0], "", *values[2:])
        stamps = []
        for index in ((0, 1, 2) if kind == "presentation" else (0, 2)):
            try:
                stat = Path(values[index]).stat()
                stamps.append((stat.st_size, stat.st_mtime_ns))
            except OSError:
                stamps.append(None)
        return values + (tuple(stamps),)

    def read_schedule(self, kind):
        inputs = self.schedule_inputs(kind)
        self.remember_settings()
        try:
            first, last = date.fromisoformat(inputs[3]), date.fromisoformat(inputs[4])
            if first > last:
                raise ValueError("시작일은 종료일보다 늦을 수 없습니다.")
        except ValueError as error:
            messagebox.showerror("날짜 확인", "YYYY-MM-DD 형식을 확인하세요.\n" + str(error), parent=self)
            return
        def work():
            events, excluded = load_schedule(inputs[2], first, last, kind=kind)
            roster = Workbook(inputs[0]).roster()
            return events, excluded, len(roster)
        def done(result):
            self.events[kind], excluded, count = result
            self.generation_signatures[kind] = inputs
            tree = self.event_tables[kind]
            tree.delete(*tree.get_children())
            labels = {"class": "수업", "project": "프로젝트", "townhall": "타운홀", "career": "커리어", "orientation": "오리엔테이션", "final": "수료"}
            for index, event in enumerate(self.events[kind]):
                category = event.get("event_type", "")
                tree.insert("", "end", iid=str(index), values=(event["date"], event["event_name"], event["instructor"], labels.get(category, category)))
            tree.selection_set(tree.get_children())
            detail = f" · PPT {4 + count * 4}장" if kind == "presentation" else ""
            self.generate_summaries[kind].set(f"명단 {count}명{detail} · 생성 대상 {len(self.events[kind])}건 · 제외 {len(excluded)}건. 수정한 양식을 기준으로 생성합니다.")
            self.status.set("일정을 불러왔습니다. 생성할 행을 선택하세요. Ctrl·Shift로 여러 행을 선택할 수 있습니다.")
        self.run("교육 일정과 출석부 명단을 읽고 있습니다…", work, done)

    def edit_instructor(self, event, kind):
        if self.busy:
            return
        tree = self.event_tables[kind]
        item = tree.identify_row(event.y)
        if not item:
            return
        record = self.events[kind][int(item)]
        answer = simpledialog.askstring("강사명 수정", record["date"] + " " + record["event_name"], initialvalue=record["instructor"], parent=self)
        if answer is not None:
            if kind == "presentation" and not presentation_eligible({**record, "instructor": answer}):
                messagebox.showinfo("강사 확인", "산출물은 강사가 지정된 수업만 생성합니다. 강사명을 입력하세요.", parent=self)
                return
            record["instructor"] = answer.strip() or "강사미정"
            tree.set(item, "instructor", record["instructor"])

    def generate(self, kind):
        if self.generation_signatures.get(kind) != self.schedule_inputs(kind):
            messagebox.showinfo("일정 확인", "양식·명단·일정 또는 기간이 바뀌었거나 아직 불러오지 않았습니다. ‘일정 불러오기’를 누르세요.", parent=self)
            return
        selected = [dict(self.events[kind][int(i)]) for i in self.event_tables[kind].selection()]
        if not selected:
            messagebox.showinfo("일정 선택", "생성할 일정 행을 선택하세요.", parent=self)
            return
        output = self.output_folder()
        if not output:
            return
        template = self.value("attendance_template")
        ppt = self.value("ppt_template") if kind == "presentation" else None
        make_zip = self.zip_flags[kind].get()
        def done(result):
            self.last_folder = result["folder"]
            self.status.set(f"{len(result['dates'])}일 · {len(result['files'])}개 파일 생성 완료: {result['folder']}")
            messagebox.showinfo("생성 완료", f"{len(result['dates'])}일, {len(result['files'])}개 파일을 만들었습니다.\n\n{result['folder']}", parent=self)
        self.run("선택한 종류의 파일을 생성하고 있습니다…", lambda: generate_files(template, ppt, selected, output, make_zip, kind=kind), done)

    def open_folder(self):
        if self.last_folder and Path(self.last_folder).is_dir():
            os.startfile(self.last_folder)
        else:
            messagebox.showinfo("결과 폴더", "먼저 결과 파일을 저장하거나 생성하세요.", parent=self)

    def close(self):
        if self.busy:
            messagebox.showinfo("처리 중", "파일 처리가 끝난 뒤 종료하세요.", parent=self)
            return
        self.remember_settings()
        self.destroy()


def main():
    app = Application()
    app.mainloop()
