import keyboard
import cv2
import numpy as np
import sys
import win32gui, win32con, win32ui, win32api
import time
import os
from pywinauto import mouse


def start_program():
    global running
    running = True
    print("Program is running")


def stop_program():
    global running
    running = False
    print("Program is stopped")


def terminate_program():
    print("Program is terminated")
    sys.exit()


# 이미지 캡처#
def take_screenshot(hwnd):
    rect = win32gui.GetClientRect(hwnd)
    width, height = rect[2] - rect[0], rect[3] - rect[1]
    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()
    save_bitmap = win32ui.CreateBitmap()
    save_bitmap.CreateCompatibleBitmap(mfc_dc, width, height)
    save_dc.SelectObject(save_bitmap)
    save_dc.BitBlt((0, 0), (width, height), mfc_dc, (0, 0), win32con.SRCCOPY)
    signed_ints_array = save_bitmap.GetBitmapBits(True)
    img = np.frombuffer(signed_ints_array, dtype="uint8")
    img.shape = (height, width, 4)
    img = img[:, :, :3]
    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwnd_dc)
    win32gui.DeleteObject(save_bitmap.GetHandle())

    return img


# 이미지 찾기#
def find_image_with_scaling(
    screenshot, template_path, threshold, scale_range=(0.5, 2.0), scale_step=0.1
):
    template = cv2.imread(template_path, cv2.IMREAD_GRAYSCALE)

    for scale in np.arange(scale_range[0], scale_range[1], scale_step):
        resized_screenshot = cv2.resize(
            screenshot, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA
        )
        result = cv2.matchTemplate(resized_screenshot, template, cv2.TM_CCOEFF_NORMED)
        y, x = np.where(result >= threshold)

        if len(x) > 0:
            return True

    return False


# 이미지 클릭 #
def click_image_with_scaling(
    hwnd, screenshot, template_path, threshold, scale_range=(0.5, 2.0), scale_step=0.1
):
    template = cv2.imread(template_path, cv2.IMREAD_GRAYSCALE)

    for scale in np.arange(scale_range[0], scale_range[1], scale_step):
        resized_screenshot = cv2.resize(
            screenshot, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA
        )
        result = cv2.matchTemplate(resized_screenshot, template, cv2.TM_CCOEFF_NORMED)
        y, x = np.where(result >= threshold)

        if len(x) > 0:
            rect = win32gui.GetClientRect(hwnd)
            x_offset = int(template.shape[1] / 2)  # 템플릿 이미지의 너비의 절반
            y_offset = int(template.shape[0] / 2)  # 템플릿 이미지의 높이의 절반
            x_coord = int(rect[0] + x[0] / scale + x_offset)
            y_coord = int(rect[1] + y[0] / scale + y_offset)
            mouse.click(button="left", coords=(x_coord, y_coord))

            return True
    return False


odin_title = "ODIN  "
hwnd = win32gui.FindWindow(None, odin_title)

if not hwnd:
    print(f"Cannot find window with title {odin_title}")
    exit()

running = False
keyboard.add_hotkey("F6", start_program)
keyboard.add_hotkey("F7", stop_program)
keyboard.add_hotkey("F8", terminate_program)

while True:
    time.sleep(0.1)

    if running:
        screenshot = take_screenshot(hwnd)
        if click_image_with_scaling(hwnd, screenshot, "quest_start.png", 0.9):
            print("퀘스트 시작 버튼 을 찾았습니다.")
            time.sleep(2)
            screenshot = take_screenshot(hwnd)
            if click_image_with_scaling(hwnd, screenshot, "quest_start2.png", 0.7):
                print("퀘스트 시작 확인 을 찾았습니다.")
                time.sleep(5)
        else:
            print("퀘스트 시작 버튼을 못찾았습니다.")

        time.sleep(1)
        working = True

        while working:
            if not running:
                break

            screenshot = take_screenshot(hwnd)

            if find_image_with_scaling(screenshot, "working.png", 0.7):
                print("이동중인 이미지를 찾았습니다. 대기합니다.")
                time.sleep(1)
            else:
                print("이동중인 이미지를 못찾았습니다.")
                break

            time.sleep(1)

        if not running:
            continue

        screenshot = take_screenshot(hwnd)

        if click_image_with_scaling(hwnd, screenshot, "quest_complete.png", 0.8):
            print("퀘스트 완료 버튼 을 찾았습니다.")
        else:
            print("퀘스트 완료 버튼을 못찾았습니다.")

        if click_image_with_scaling(hwnd, screenshot, "rewards.png", 0.8):
            print("보상 버튼 을 찾았습니다.")
        else:
            print("보상 버튼을 못찾았습니다.")

        time.sleep(1)

        if click_image_with_scaling(hwnd, screenshot, "skip.png", 0.8):
            print("스킵 버튼 을 찾았습니다.")
        else:
            print("스킵 버튼을 못찾았습니다.")

        time.sleep(1)

        Talk = True
        while Talk:
            if not running:
                break

            screenshot = take_screenshot(hwnd)

            if click_image_with_scaling(hwnd, screenshot, "talk.png", 0.8):
                print("대화 버튼 을 찾았습니다")
            else:
                print("대화 버튼을 못 찾았습니다.")
                break

            time.sleep(1)

        potion = True
        while potion:
            if not running:
                break

            screenshot = take_screenshot(hwnd)

            if find_image_with_scaling(screenshot, "potion2.png", 0.7):
                print("포션이 다 떨어 젓습니다.")
                time.sleep(1)
                if click_image_with_scaling(hwnd, screenshot, "town.png", 0.7):
                    print("마을로 이동합니다.")
                    time.sleep(2)
                    screenshot = take_screenshot(hwnd)
                    if click_image_with_scaling(hwnd, screenshot, "ok.png", 0.9):
                        print("확인 버튼을 누릅니다.")
                        time.sleep(8)
                        screenshot = take_screenshot(hwnd)
                        if click_image_with_scaling(
                            hwnd, screenshot, "Potion Shop.png", 0.8
                        ):
                            print("포션 상점으로 이동합니다.")
                            time.sleep(13)
                            screenshot = take_screenshot(hwnd)
                            if click_image_with_scaling(
                                hwnd, screenshot, "potion buy.png", 0.8
                            ):
                                print("포션을 구매합니다.")
                                time.sleep(1)
                                if click_image_with_scaling(
                                    hwnd, screenshot, "out.png", 0.8
                                ):
                                    print("상점을 나갑니다.")
                                    time.sleep(1)
            else:
                print("포션이 충분합니다.")
                break

            time.sleep(1)

    # 종료 조건을 확인합니다.
    if keyboard.is_pressed("F8"):
        terminate_program()
