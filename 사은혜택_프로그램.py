# -*- coding: utf-8 -*-
"""
================================================================================
 AK PLAZA 카드 사은혜택 자동 입력 - 마우스로 쓰는 창(GUI) 버전
================================================================================
 - 윈도우 계산기처럼 창을 띄워 마우스로 클릭해서 사용합니다. (오프라인 동작)
 - PPT/PDF 파일과 엑셀 양식은 창 안으로 '드래그앤드롭'하거나
   '파일 찾아보기' 버튼으로 선택할 수 있습니다.
 - 실제 처리는 update_calendar.py 의 함수를 그대로 사용합니다.

 실행: '프로그램 실행.bat' 을 더블클릭하세요.
================================================================================
"""

import os
import sys
import queue
import threading
import traceback

# 콘솔이 없어도(창 모드) 안전하도록 표준출력 인코딩 설정
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# 같은 폴더의 update_calendar.py 를 불러옵니다.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import update_calendar as core   # noqa: E402

import tkinter as tk              # noqa: E402
from tkinter import ttk, filedialog, messagebox  # noqa: E402

# 드래그앤드롭(선택). 없으면 '파일 찾아보기' 버튼만 사용합니다.
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_사용가능 = True
except Exception:
    DND_사용가능 = False


# ------------------------------------------------------------------------------
# 색/글꼴
# ------------------------------------------------------------------------------
배경 = "#f4f6f8"
강조 = "#0b5cab"
연회색 = "#e9eef4"
글꼴 = ("맑은 고딕", 10)
글꼴_굵게 = ("맑은 고딕", 11, "bold")
글꼴_제목 = ("맑은 고딕", 15, "bold")


def 경로정리(raw):
    """드롭/선택된 문자열에서 실제 파일 경로 하나를 뽑아냅니다.
    (드래그앤드롭은 '{C:\\경로 있는 파일.pptx}' 처럼 중괄호로 올 수 있습니다.
     공백이 들어간 폴더 경로도 잘리지 않도록 처리합니다.)"""
    if not raw:
        return ""
    raw = raw.strip()
    if raw.startswith("{") and "}" in raw:          # 중괄호로 감싸진 경우
        return raw[1:raw.index("}")].strip()
    raw = raw.strip('"')
    if os.path.exists(raw):                          # 공백 포함이라도 실제 파일이면 그대로
        return raw
    return (raw.split()[0] if raw else raw).strip().strip('"')   # 여러 파일 중 첫 번째


class 앱:
    def __init__(self, root):
        self.root = root
        root.title("AK PLAZA 카드 사은혜택 자동 입력")
        root.configure(bg=배경)
        try:
            root.geometry("640x620")
            root.minsize(560, 560)
        except Exception:
            pass

        self.양식_경로 = tk.StringVar()
        self.입력_경로 = tk.StringVar()
        self.발급포함 = tk.BooleanVar(value=True)
        self.정기포함 = tk.BooleanVar(value=True)
        self.마지막결과폴더 = None
        self.작업큐 = queue.Queue()   # 워커 스레드 → 메인스레드 전달용

        self._헤더()
        self._파일칸("① 엑셀 양식 파일 (.xlsx)", self.양식_경로,
                   [("엑셀 파일", "*.xlsx")], "xlsx")
        self._파일칸("② 이번 달 캘린더 파일 (PPT 권장 / PDF 가능)", self.입력_경로,
                   [("PPT/PDF", "*.pptx *.pdf"), ("PPT", "*.pptx"), ("PDF", "*.pdf")], "input")
        self._옵션()
        self._실행버튼()
        self._로그()

        self._자동채우기()
        self.root.after(150, self._큐처리)   # 메인스레드에서 주기적으로 큐 확인

    # ----- 화면 구성 -----
    def _헤더(self):
        f = tk.Frame(self.root, bg=강조)
        f.pack(fill="x")
        tk.Label(f, text="카드 사은혜택 자동 입력", font=글꼴_제목,
                 bg=강조, fg="white").pack(anchor="w", padx=16, pady=(12, 2))
        안내 = "파일을 아래 칸에 끌어다 놓거나(드래그앤드롭), '파일 찾아보기'로 선택하세요."
        if not DND_사용가능:
            안내 = "'파일 찾아보기' 버튼으로 파일을 선택하세요."
        tk.Label(f, text=안내, font=글꼴, bg=강조, fg="#dce8f5").pack(anchor="w", padx=16, pady=(0, 12))

    def _파일칸(self, 제목, 변수, 필터, 종류):
        box = tk.LabelFrame(self.root, text=" " + 제목 + " ", font=글꼴_굵게,
                            bg=배경, fg="#222", padx=10, pady=8)
        box.pack(fill="x", padx=14, pady=(12, 0))

        drop = tk.Label(box,
                        text=("여기로 파일을 끌어다 놓으세요" if DND_사용가능 else "아래 버튼으로 파일을 선택하세요"),
                        bg="white", fg="#888", font=글꼴, relief="solid", bd=1,
                        height=2, anchor="center")
        drop.pack(fill="x", pady=(2, 6))
        if DND_사용가능:
            drop.drop_target_register(DND_FILES)
            drop.dnd_bind("<<Drop>>", lambda e, v=변수, k=종류, d=drop: self._드롭(e, v, k, d))

        row = tk.Frame(box, bg=배경)
        row.pack(fill="x")
        ent = tk.Entry(row, textvariable=변수, font=글꼴)
        ent.pack(side="left", fill="x", expand=True, ipady=3)
        tk.Button(row, text="파일 찾아보기", font=글꼴, bg=연회색,
                  command=lambda v=변수, ft=필터, k=종류, d=drop: self._찾아보기(v, ft, k, d)
                  ).pack(side="left", padx=(8, 0))
        # 표시용 드롭 라벨 저장
        setattr(self, "drop_" + 종류, drop)

    def _옵션(self):
        box = tk.LabelFrame(self.root, text=" ③ 포함할 항목 선택 ", font=글꼴_굵게,
                            bg=배경, fg="#222", padx=10, pady=6)
        box.pack(fill="x", padx=14, pady=(12, 0))
        tk.Checkbutton(box, text="신한Plus 발급①/발급② 혜택 포함", variable=self.발급포함,
                       bg=배경, font=글꼴, anchor="w").pack(fill="x")
        tk.Checkbutton(box, text="신한Plus 정기 혜택 포함", variable=self.정기포함,
                       bg=배경, font=글꼴, anchor="w").pack(fill="x")

    def _실행버튼(self):
        f = tk.Frame(self.root, bg=배경)
        f.pack(fill="x", padx=14, pady=12)
        self.실행btn = tk.Button(f, text="▶  엑셀 만들기", font=("맑은 고딕", 12, "bold"),
                                bg=강조, fg="white", activebackground="#094a8c",
                                activeforeground="white", relief="flat", height=2,
                                command=self._실행클릭)
        self.실행btn.pack(side="left", fill="x", expand=True)
        self.열기btn = tk.Button(f, text="📂 결과 폴더 열기", font=글꼴, bg=연회색,
                                state="disabled", command=self._폴더열기)
        self.열기btn.pack(side="left", padx=(8, 0), fill="y")

    def _로그(self):
        box = tk.LabelFrame(self.root, text=" 진행 상황 ", font=글꼴_굵게,
                            bg=배경, fg="#222", padx=6, pady=6)
        box.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        self.로그창 = tk.Text(box, height=8, font=("맑은 고딕", 9), bg="#20262e",
                            fg="#e6edf3", relief="flat", wrap="word")
        self.로그창.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(box, command=self.로그창.yview)
        sb.pack(side="right", fill="y")
        self.로그창.config(yscrollcommand=sb.set, state="disabled")

    # ----- 동작 -----
    def _로그출력(self, 글):
        """워커 스레드에서 호출됨 → 큐에만 넣고, 실제 화면 갱신은 메인스레드가 함."""
        self.작업큐.put(("log", str(글)))

    def _화면에로그(self, 글):
        self.로그창.config(state="normal")
        self.로그창.insert("end", str(글) + "\n")
        self.로그창.see("end")
        self.로그창.config(state="disabled")

    def _큐처리(self):
        """메인스레드에서 주기적으로 큐를 비우며 화면을 갱신합니다(Tk는 메인스레드에서만 조작)."""
        try:
            while True:
                종류, 값 = self.작업큐.get_nowait()
                if 종류 == "log":
                    self._화면에로그(값)
                elif 종류 == "done":
                    결과, 폴더 = 값
                    self.마지막결과폴더 = 폴더
                    self.실행btn.config(state="normal", text="▶  엑셀 만들기")
                    self.열기btn.config(state="normal")
                    이름 = os.path.basename(결과.get("결과경로", ""))
                    messagebox.showinfo("완료",
                                        "엑셀 파일을 만들었습니다!\n\n" + 이름 +
                                        "\n\n'결과 폴더 열기' 버튼으로 확인하세요.")
                elif 종류 == "error":
                    메시지, 자세히 = 값
                    self._화면에로그("\n[오류] " + 메시지)
                    self._화면에로그(자세히)
                    self.실행btn.config(state="normal", text="▶  엑셀 만들기")
                    messagebox.showerror("오류", "처리 중 문제가 발생했습니다.\n\n" + 메시지)
        except queue.Empty:
            pass
        self.root.after(150, self._큐처리)

    def _드롭(self, event, 변수, 종류, drop라벨):
        경로 = 경로정리(event.data)
        self._설정경로(변수, 경로, 종류, drop라벨)

    def _찾아보기(self, 변수, 필터, 종류, drop라벨):
        경로 = filedialog.askopenfilename(filetypes=필터 + [("모든 파일", "*.*")])
        if 경로:
            self._설정경로(변수, 경로, 종류, drop라벨)

    def _설정경로(self, 변수, 경로, 종류, drop라벨):
        경로 = 경로정리(경로)
        if not 경로 or not os.path.exists(경로):
            return
        if 종류 == "input" and core.입력종류_판단(경로) is None:
            messagebox.showwarning("파일 형식", "PPT(.pptx) 또는 PDF(.pdf) 파일만 넣어 주세요.")
            return
        if 종류 == "xlsx" and not 경로.lower().endswith(".xlsx"):
            messagebox.showwarning("파일 형식", "엑셀(.xlsx) 파일을 넣어 주세요.")
            return
        변수.set(경로)
        drop라벨.config(text="✔  " + os.path.basename(경로), fg="#0b7a3b")

    def _자동채우기(self):
        """프로그램 폴더에서 양식/입력 파일을 자동으로 찾아 미리 넣어 둡니다."""
        폴더 = os.path.dirname(os.path.abspath(__file__))
        try:
            x, p, d = core.파일_자동찾기(폴더)
        except Exception:
            x = p = d = None
        if x:
            self.양식_경로.set(x)
            self.drop_xlsx.config(text="✔  " + os.path.basename(x), fg="#0b7a3b")
        inp = p or d
        if inp:
            self.입력_경로.set(inp)
            self.drop_input.config(text="✔  " + os.path.basename(inp), fg="#0b7a3b")

    def _실행클릭(self):
        양식 = self.양식_경로.get().strip()
        입력 = self.입력_경로.get().strip()
        if not 양식 or not os.path.exists(양식):
            messagebox.showwarning("확인", "엑셀 양식 파일(.xlsx)을 먼저 넣어 주세요.")
            return
        if not 입력 or not os.path.exists(입력):
            messagebox.showwarning("확인", "이번 달 PPT(또는 PDF) 파일을 먼저 넣어 주세요.")
            return
        종류 = core.입력종류_판단(입력)
        없음 = core.라이브러리_확인(종류)
        if 없음:
            messagebox.showerror("부품 필요",
                                 "다음 부품이 설치되어 있지 않습니다:\n" + ", ".join(없음) +
                                 "\n\n먼저 '최초설치.bat' 을 한 번 실행해 주세요.")
            return
        # 로그 비우기
        self.로그창.config(state="normal"); self.로그창.delete("1.0", "end")
        self.로그창.config(state="disabled")
        self.실행btn.config(state="disabled", text="처리 중...")
        self.열기btn.config(state="disabled")

        폴더 = os.path.dirname(os.path.abspath(입력))   # 결과는 입력 파일 폴더에 저장

        발급 = self.발급포함.get()
        정기 = self.정기포함.get()

        def 작업():
            try:
                결과 = core.혜택_처리(양식, 입력, 종류, 발급, 정기, 폴더, 로그=self._로그출력)
                self._로그출력("\n=== 완료! '결과 폴더 열기' 버튼으로 확인하세요. ===")
                self.작업큐.put(("done", (결과, 폴더)))
            except Exception as e:
                self.작업큐.put(("error", (str(e), traceback.format_exc())))
        threading.Thread(target=작업, daemon=True).start()

    def _폴더열기(self):
        폴더 = self.마지막결과폴더
        if not 폴더 or not os.path.isdir(폴더):
            return
        try:
            os.startfile(폴더)                       # 윈도우
        except AttributeError:
            import subprocess
            opener = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.Popen([opener, 폴더])


def main():
    root = TkinterDnD.Tk() if DND_사용가능 else tk.Tk()
    앱(root)
    root.mainloop()


if __name__ == "__main__":
    main()
