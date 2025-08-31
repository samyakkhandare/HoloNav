import cv2, mediapipe as mp, pyautogui, threading, speech_recognition as sr
import subprocess, webbrowser, pygetwindow as gw, os, pywhatkit as kit

# ------------------ Apps & Files ------------------
apps = {
    "youtube": "https://www.youtube.com",
    "notepad": "notepad.exe",
    "calculator": "calc.exe"
}
ppt_path = r"C:\Users\intel\OneDrive\Desktop\PPT OF MY.pptx

# ------------------ Voice Control ------------------
def voice_listener():
    recognizer = sr.Recognizer()
    mic = sr.Microphone()
    youtube_open = False

    while True:
        try:
            with mic as source:
                print("🎤 Listening...")
                recognizer.adjust_for_ambient_noise(source, duration=1)
                audio = recognizer.listen(source)

            command = recognizer.recognize_google(audio).lower()
            print(f"👉 Command: {command}")

            # --- Apps ---
            if "powerpoint" in command:
                subprocess.Popen(["start", "powerpnt"], shell=True)
            elif "open presentation" in command or "open ppt" in command:
                os.startfile(ppt_path)
            elif "youtube" in command:
                webbrowser.open(apps["youtube"]); youtube_open = True
            elif "close youtube" in command and youtube_open:
                pyautogui.hotkey('alt', 'f4'); youtube_open = False
            elif "notepad" in command:
                subprocess.Popen(apps["notepad"])
            elif "calculator" in command:
                subprocess.Popen(apps["calculator"])

            # --- Simple Direct Play like YouTube mic ---
            elif "play" in command:
                search = command.replace("play", "").strip()
                if search == "":
                    search = "latest songs"
                print(f"🎶 Playing: {search}")
                kit.playonyt(search)   # ✅ Directly auto-play first result
                youtube_open = True

            # --- YouTube Controls ---
            elif youtube_open and ("pause" in command or "band" in command or "ruk" in command):
                pyautogui.press("k")   # pause/play toggle
            elif youtube_open and ("resume" in command or "start" in command or "laga do" in command):
                pyautogui.press("k")   # resume
            elif "exit" in command or "band kar" in command:
                if youtube_open:
                    pyautogui.hotkey('alt', 'f4')
                    youtube_open = False
                print("👋 Exiting program...")
                os._exit(0)

        except sr.UnknownValueError:
            continue
        except sr.RequestError:
            print("⚠️ Speech API error")
            break

# Run voice listener in background
threading.Thread(target=voice_listener, daemon=True).start()

# ------------------ Gesture Control ------------------
pose = mp.solutions.pose.Pose()
draw = mp.solutions.drawing_utils
cap = cv2.VideoCapture(0)

def active_window():
    try:
        win = gw.getActiveWindow()
        return win.title.lower() if win else ""
    except:
        return ""

while True:
    ret, frame = cap.read()
    if not ret: break
    frame = cv2.flip(frame, 1)
    rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    results = pose.process(rgb)

    if results.pose_landmarks:
        lm = results.pose_landmarks.landmark
        left, right = lm[mp.solutions.pose.PoseLandmark.LEFT_WRIST], lm[mp.solutions.pose.PoseLandmark.RIGHT_WRIST]
        l_sh, r_sh = lm[mp.solutions.pose.PoseLandmark.LEFT_SHOULDER], lm[mp.solutions.pose.PoseLandmark.RIGHT_SHOULDER]
        left_up, right_up = left.y < l_sh.y - 0.05, right.y < r_sh.y - 0.05
        win = active_window()

        # --- PowerPoint Controls ---
        if "powerpoint" in win:
            if right_up and not left_up: pyautogui.press("right")   # next slide
            elif left_up and not right_up: pyautogui.press("left")  # previous slide
            elif left_up and right_up: pyautogui.press("f5")        # slideshow

        # --- YouTube / Chrome Controls ---
        elif "youtube" in win or "chrome" in win:
            if right_up and not left_up: pyautogui.press("l")   # forward 10s
            elif left_up and not right_up: pyautogui.press("j") # backward 10s
            elif left_up and right_up: pyautogui.press("space") # play/pause

        draw.draw_landmarks(frame, results.pose_landmarks, mp.solutions.pose.POSE_CONNECTIONS)

    cv2.imshow("🎥 Voice + Gesture Control", frame)
    if cv2.waitKey(1) & 0xFF == ord('q'): break

cap.release(); cv2.destroyAllWindows()
