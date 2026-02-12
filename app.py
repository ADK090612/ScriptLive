import sys
import json
import os
import tempfile
import time
import threading
from datetime import datetime, timedelta

from PyQt6.QtWidgets import (QApplication, QMainWindow, QMessageBox)
from PyQt6.QtCore import QTimer, QTime, Qt
from PyQt6 import uic

import pyttsx3

# [수정됨] sounddevice 로드 실패 시(예: ARM64) 프로그램 뻗음 방지 및 대체 수단 마련
HAS_SOUNDDEVICE = False
try:
    import sounddevice as sd
    import soundfile as sf
    import numpy as np
    HAS_SOUNDDEVICE = True
except Exception as e:
    print(f"Warning: sounddevice 라이브러리를 로드할 수 없습니다. ({e})")
    HAS_SOUNDDEVICE = False

# 윈도우 기본 사운드 재생 (Fallback 용)
import winsound

# 템플릿 저장 파일명
TEMPLATE_FILE = "templates.json"

class AnnouncerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        # .ui 파일 로드
        uic.loadUi("announcer.ui", self)

        # TTS 엔진 초기화
        self.engine = pyttsx3.init()
        
        # 상태 변수
        self.is_scheduled = False
        self.broadcast_timer = QTimer(self)
        self.broadcast_timer.timeout.connect(self.check_schedule)
        self.target_time = None
        self.voices = []
        self.audio_devices = []

        # 초기화 함수 호출
        self.init_ui()
        self.load_voices()
        self.load_audio_devices() # 여기서 실패 시 경고창 띄움
        self.load_templates()

    def init_ui(self):
        """UI 이벤트 연결 및 초기 설정"""
        self.time_edit.setTime(QTime.currentTime())
        
        # 버튼 연결
        self.btn_preview_voice.clicked.connect(self.preview_voice)
        self.btn_preview_script.clicked.connect(self.preview_script)
        self.btn_schedule.clicked.connect(self.start_schedule)
        self.btn_stop.clicked.connect(self.reset_state)
        self.btn_save_tmpl.clicked.connect(self.save_template)
        self.btn_del_tmpl.clicked.connect(self.delete_template)
        
        # 템플릿 선택 시 로드
        self.combo_template.currentIndexChanged.connect(self.load_selected_template)

    def load_voices(self):
        """설치된 TTS 음성 목록 로드"""
        self.voices = self.engine.getProperty('voices')
        self.combo_voice.clear()
        
        for v in self.voices:
            # 한국어 우선 정렬 로직 등은 생략하고 단순 추가
            self.combo_voice.addItem(f"{v.name} ({v.languages})")
        
        # 한국어 음성이 있다면 기본 선택 시도
        for i, v in enumerate(self.voices):
            if 'KR' in v.id or 'Korean' in v.name:
                self.combo_voice.setCurrentIndex(i)
                break

    def load_audio_devices(self):
        """시스템 오디오 출력 장치 로드 (sounddevice) - 방어 코드 추가됨"""
        self.audio_devices = [] 
        self.combo_output.clear()

        # [수정됨] sounddevice 모듈 자체가 없을 경우 (Import Error 등)
        if not HAS_SOUNDDEVICE:
            QMessageBox.warning(self, "오디오 장치 경고", 
                                "고급 오디오 제어(sounddevice)를 사용할 수 없습니다.\n"
                                "윈도우 기본 재생 장치(Winsound)를 사용합니다.")
            self.combo_output.addItem("시스템 기본 장치 (Winsound)")
            return

        try:
            # [수정됨] 장치 쿼리 중 에러 발생 시 방어
            devices = sd.query_devices()
            
            # 출력 장치만 필터링 (> 0 output channels)
            for i, d in enumerate(devices):
                if d['max_output_channels'] > 0:
                    name = f"{d['name']} ({d['hostapi']})"
                    self.combo_output.addItem(name)
                    self.audio_devices.append(i) # 실제 device ID 저장
            
            # 기본 장치 선택 (시스템 기본값)
            try:
                default_device = sd.default.device[1]
                if default_device in self.audio_devices:
                    idx = self.audio_devices.index(default_device)
                    self.combo_output.setCurrentIndex(idx)
            except:
                pass # 기본 장치 선택 실패 시 0번 유지
                
        except Exception as e:
            # [수정됨] 런타임 에러 발생 시 처리
            self.label_status.setText(f"오디오 장치 로드 실패: {e}")
            self.combo_output.addItem("시스템 기본 장치 (Winsound Fallback)")
            # self.audio_devices 리스트가 비어있으므로 자동으로 Fallback 로직을 타게 됨

    def load_templates(self):
        """JSON에서 템플릿 불러오기"""
        self.combo_template.clear()
        self.combo_template.addItem("-- 저장된 대본 선택 --")
        
        if os.path.exists(TEMPLATE_FILE):
            with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
                self.templates = json.load(f)
                for tmpl in self.templates:
                    self.combo_template.addItem(tmpl['title'])
        else:
            self.templates = []

    def save_template(self):
        title = self.input_title.text().strip()
        script = self.text_script.toPlainText().strip()
        
        if not title or not script:
            QMessageBox.warning(self, "경고", "제목과 대본을 모두 입력해주세요.")
            return

        self.templates.append({"title": title, "script": script})
        with open(TEMPLATE_FILE, "w", encoding="utf-8") as f:
            json.dump(self.templates, f, ensure_ascii=False, indent=4)
        
        self.load_templates()
        self.combo_template.setCurrentIndex(len(self.templates)) # 방금 저장한 것 선택
        self.label_status.setText("템플릿 저장 완료!")

    def delete_template(self):
        idx = self.combo_template.currentIndex()
        if idx <= 0: return # 기본 문구 선택 시 무시

        if QMessageBox.question(self, "삭제", "정말 삭제하시겠습니까?") == QMessageBox.StandardButton.Yes:
            del self.templates[idx - 1]
            with open(TEMPLATE_FILE, "w", encoding="utf-8") as f:
                json.dump(self.templates, f, ensure_ascii=False, indent=4)
            
            self.load_templates()
            self.input_title.clear()
            self.text_script.clear()

    def load_selected_template(self, index):
        if index <= 0: return
        data = self.templates[index - 1]
        self.input_title.setText(data['title'])
        self.text_script.setText(data['script'])

    def get_selected_voice_id(self):
        idx = self.combo_voice.currentIndex()
        if idx >= 0:
            return self.voices[idx].id
        return None

    def get_selected_output_device(self):
        """콤보박스 인덱스를 sounddevice ID로 변환"""
        # [수정됨] sounddevice가 없거나 목록이 비어있으면 None 반환
        if not HAS_SOUNDDEVICE or not self.audio_devices:
            return None
            
        idx = self.combo_output.currentIndex()
        if idx >= 0 and idx < len(self.audio_devices):
            return self.audio_devices[idx]
        return None

    def generate_and_play(self, text):
        """
        핵심 기능: TTS -> WAV 파일 -> 특정 오디오 장치로 재생
        이 과정은 UI 멈춤 방지를 위해 별도 스레드에서 실행하는 것이 좋으나,
        간단한 구현을 위해 여기서는 함수 내에서 처리합니다.
        """
        voice_id = self.get_selected_voice_id()
        device_id = self.get_selected_output_device()
        
        if not voice_id: return

        # 스레드에서 재생 (UI 프리징 방지)
        t = threading.Thread(target=self._play_thread, args=(text, voice_id, device_id))
        t.start()

    def _play_thread(self, text, voice_id, device_id):
        """실제 재생 로직 (백그라운드) - winsound fallback 추가"""
        temp_path = ""
        try:
            # 1. 임시 파일 생성
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, "tts_temp.wav")
            
            # 2. TTS 엔진을 사용하여 파일로 저장 (공통)
            engine = pyttsx3.init()
            engine.setProperty('voice', voice_id)
            engine.save_to_file(text, temp_path)
            engine.runAndWait()
            
            if not os.path.exists(temp_path):
                return

            # [수정됨] 3. SoundDevice 사용 가능 여부에 따른 분기 처리
            if HAS_SOUNDDEVICE and device_id is not None:
                # A. SoundDevice로 특정 장치 재생
                data, fs = sf.read(temp_path)
                self.update_status(f"재생 중... (장치 ID: {device_id})", True)
                sd.play(data, fs, device=device_id, blocking=True)
            else:
                # B. Winsound로 기본 장치 재생 (Fallback)
                self.update_status("재생 중... (기본 장치)", True)
                winsound.PlaySound(temp_path, winsound.SND_FILENAME)

            self.update_status("대기 중...", False)
            
        except Exception as e:
            print(f"Error: {e}")
            self.update_status(f"오류 발생: {e}", False)
        finally:
            # 파일 삭제 시도
            try:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
            except:
                pass

    def update_status(self, text, is_playing):
        """스레드 안전하게 UI 업데이트"""
        self.label_status.setText(text)
        if is_playing:
            self.label_status.setStyleSheet("color: #00e676; font-weight: bold;")
        else:
            self.label_status.setStyleSheet("color: #ff4081;")

    def preview_voice(self):
        self.generate_and_play("안녕하세요? 음성 모델 테스트 중입니다.")

    def preview_script(self):
        script = self.text_script.toPlainText()
        if not script: return
        self.generate_and_play(script)

    def start_schedule(self):
        if self.is_scheduled: return

        script = self.text_script.toPlainText()
        if not script:
            QMessageBox.warning(self, "알림", "스크립트를 입력하세요.")
            return

        # 시간 계산
        now = datetime.now()
        
        if self.tabWidget.currentIndex() == 0: # 특정 시간
            qtime = self.time_edit.time()
            target = datetime(now.year, now.month, now.day, qtime.hour(), qtime.minute(), qtime.second())
            if target <= now:
                target += timedelta(days=1) # 내일로 설정
        else: # 타이머
            h = self.spin_h.value()
            m = self.spin_m.value()
            s = self.spin_s.value()
            if h == 0 and m == 0 and s == 0:
                QMessageBox.warning(self, "경고", "타이머 시간을 설정하세요.")
                return
            target = now + timedelta(hours=h, minutes=m, seconds=s)

        self.target_time = target
        self.is_scheduled = True
        self.btn_schedule.setEnabled(False)
        self.btn_schedule.setStyleSheet("background-color: gray; color: black;")
        self.broadcast_timer.start(1000) # 1초마다 체크

    def check_schedule(self):
        if not self.is_scheduled: return

        now = datetime.now()
        remaining = self.target_time - now
        
        if remaining.total_seconds() <= 0:
            self.broadcast_timer.stop()
            self.label_status.setText("ON AIR - 방송 송출 중")
            script = self.text_script.toPlainText()
            self.generate_and_play(script)
            
            # 방송 후 리셋
            # 재생 시간만큼 기다려야 하지만, 여기서는 즉시 리셋 상태로 복귀 (재생은 백그라운드)
            self.reset_state()
            self.label_status.setText("방송 송출 완료")
        else:
            # 남은 시간 표시
            self.label_status.setText(f"방송 대기 중... {str(remaining).split('.')[0]}")

    def reset_state(self):
        self.is_scheduled = False
        self.broadcast_timer.stop()
        self.btn_schedule.setEnabled(True)
        self.btn_schedule.setStyleSheet("background-color: rgba(0, 230, 118, 0.2); border-color: #00e676; color: white;")
        self.label_status.setText("방송 취소됨 / 대기 중")
        
        # [수정됨] sounddevice가 있을 때만 stop 호출 (없으면 winsound stop)
        if HAS_SOUNDDEVICE:
            sd.stop()
        else:
            winsound.PlaySound(None, winsound.SND_PURGE) # winsound 중지 명령

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AnnouncerApp()
    window.show()
    sys.exit(app.exec())
