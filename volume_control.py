import cv2
import mediapipe as mp
import numpy as np
import math
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL


try:
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
except Exception as e:
    print("Pycaw Import Error:", e)
    exit()

try:
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_,
        CLSCTX_ALL,
        None
    )
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    volRange = volume.GetVolumeRange()
    minVol = volRange[0]  
    maxVol = volRange[1]   

    print(" Audio Connected Successfully")
except Exception as e:
    print("Audio Connection Error:", e)
    exit()

mp_hands = mp.solutions.hands
hands = mp_hands.Hands(
    min_detection_confidence=0.7,
    min_tracking_confidence=0.7
)
mp_draw = mp.solutions.drawing_utils

cap = cv2.VideoCapture(0)

volBar = 400
volPer = 0

while True:
    success, img = cap.read()
    if not success:
        print(" Camera not working")
        break

    img = cv2.flip(img, 1)
    rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    results = hands.process(rgb)

    if results.multi_hand_landmarks:
        for handLms in results.multi_hand_landmarks:
            lmList = []

            for id, lm in enumerate(handLms.landmark):
                h, w, c = img.shape
                cx, cy = int(lm.x * w), int(lm.y * h)
                lmList.append([id, cx, cy])
            if len(lmList) >= 9:
                x1, y1 = lmList[4][1], lmList[4][2]
                x2, y2 = lmList[8][1], lmList[8][2]

                cx, cy = (x1 + x2) // 2, (y1 + y2) // 2

                length = math.hypot(x2 - x1, y2 - y1)
                length = np.clip(length, 20, 200)
                vol = np.interp(length, [20, 200], [minVol, maxVol])
                volBar = np.interp(length, [20, 200], [400, 150])
                volPer = np.interp(length, [20, 200], [0, 100])
                smoothness = 3
                vol = smoothness * round(vol / smoothness)
                try:
                    volume.SetMasterVolumeLevel(vol, None)
                except:
                    pass
                cv2.circle(img, (x1, y1), 10, (255, 0, 255), cv2.FILLED)
                cv2.circle(img, (x2, y2), 10, (255, 0, 255), cv2.FILLED)
                cv2.line(img, (x1, y1), (x2, y2), (255, 0, 255), 3)
                cv2.circle(img, (cx, cy), 8, (0, 255, 0), cv2.FILLED)
                if length < 30:
                    cv2.circle(img, (cx, cy), 12, (0, 255, 0), cv2.FILLED)
            mp_draw.draw_landmarks(img, handLms, mp_hands.HAND_CONNECTIONS)
    cv2.rectangle(img, (50, 150), (85, 400), (0, 255, 0), 3)
    cv2.rectangle(img, (50, int(volBar)), (85, 400), (0, 255, 0), cv2.FILLED)
    cv2.putText(img, f'{int(volPer)} %', (40, 450),
                cv2.FONT_HERSHEY_COMPLEX, 1, (0, 255, 0), 2)
    cv2.imshow("Hand Gesture Volume Control", img)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break
cap.release()
cv2.destroyAllWindows()