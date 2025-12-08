import os
import time
import cv2
import numpy as np
import pyautogui
from PIL import Image
import pygetwindow as gw
import requests
import json
from pywinauto.application import Application
from pywinauto import keyboard
from pywinauto import Application
import pyperclip

# ----------------------------
# Configuration
# ----------------------------
ICON_PATH = "notepad_icon.png"  # Notepad desktop icon
OUTPUT_DIR = r"C:\Users\amrk6\Desktop\tjm-project"
ANNOTATED_DIR = os.path.join(OUTPUT_DIR, "annotated")
POSTS_API = "https://jsonplaceholder.typicode.com/posts"
MAX_POSTS = 10
RETRY_ATTEMPTS = 3
RETRY_DELAY = 1  # seconds

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(ANNOTATED_DIR, exist_ok=True)

# ----------------------------
# Desktop & Screenshot
# ----------------------------
def show_desktop():
    pyautogui.hotkey('win', 'd')
    time.sleep(0.5)  # wait for desktop animation

def take_screenshot(path="desktop.png"):
    show_desktop()
    screenshot = pyautogui.screenshot()
    screenshot.save(path)
    return path

# ----------------------------
# ORB Feature Matching
# ----------------------------
def find_icon_center_orb(template_path, screenshot_path, min_matches=4):
    template = cv2.imread(template_path, cv2.IMREAD_GRAYSCALE)
    screenshot = cv2.imread(screenshot_path, cv2.IMREAD_GRAYSCALE)
    if template is None or screenshot is None:
        return None

    orb = cv2.ORB_create(nfeatures=2000)
    kp1, des1 = orb.detectAndCompute(template, None)
    kp2, des2 = orb.detectAndCompute(screenshot, None)

    if des1 is None or des2 is None:
        return None

    bf = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=False)
    matches = bf.knnMatch(des1, des2, k=2)

    good_matches = []
    for m, n in matches:
        if m.distance < 0.85 * n.distance:
            good_matches.append(m)

    if len(good_matches) < min_matches:
        return None

    pts = np.float32([kp2[m.trainIdx].pt for m in good_matches])
    x_min, y_min = np.min(pts, axis=0)
    x_max, y_max = np.max(pts, axis=0)
    center_x = int((x_min + x_max) / 2)
    center_y = int((y_min + y_max) / 2)

    annotated = cv2.cvtColor(screenshot, cv2.COLOR_GRAY2BGR)
    cv2.rectangle(annotated, (int(x_min), int(y_min)), (int(x_max), int(y_max)), (0,0,255), 2)
    cv2.circle(annotated, (center_x, center_y), 10, (0,255,0), -1)
    annotated_path = os.path.join(ANNOTATED_DIR, f"annotated_orb_{int(time.time())}.png")
    cv2.imwrite(annotated_path, annotated)

    return (center_x, center_y)

def template_match_icon(template_path, screenshot_path, threshold=0.7):
    template = cv2.imread(template_path)
    screenshot = cv2.imread(screenshot_path)
    if template is None or screenshot is None:
        return None

    res = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
    if max_val >= threshold:
        center_x = max_loc[0] + template.shape[1] // 2
        center_y = max_loc[1] + template.shape[0] // 2
        cv2.rectangle(screenshot, max_loc, 
                      (max_loc[0]+template.shape[1], max_loc[1]+template.shape[0]), 
                      (0,255,0), 2)
        annotated_path = os.path.join(ANNOTATED_DIR, f"annotated_template_{int(time.time())}.png")
        cv2.imwrite(annotated_path, screenshot)
        return (center_x, center_y)
    return None

def detect_icon(template_path, screenshot_path):
    center = find_icon_center_orb(template_path, screenshot_path)
    if center:
        print("Icon detected with ORB at", center)
        return center
    center = template_match_icon(template_path, screenshot_path)
    if center:
        print("Icon detected with Template Matching at", center)
        return center
    return None

def click_icon(center):
    x, y = center
    pyautogui.moveTo(x, y, duration=0.3)
    pyautogui.doubleClick()

def wait_for_notepad():
    for _ in range(10):
        windows = gw.getWindowsWithTitle("Untitled - Notepad")
        if windows:
            return True
        time.sleep(0.5)
    return False

def type_and_save_post(post):
    content = f"Title: {post['title']}\n\n{post['body']}"
    pyautogui.write(content, interval=0.02)
    time.sleep(0.5)

    pyautogui.hotkey('ctrl', 's')  # open Save As
    time.sleep(0.5)

    # Navigate to the target folder
    
    full_path = os.path.join(OUTPUT_DIR, f"post_{post['id']}.txt")
    pyautogui.write(full_path)
    pyautogui.press('enter')
    time.sleep(0.5)

    pyautogui.hotkey('ctrl', 'w')  # close Notepad


# ----------------------------
# Chrome Fallback
# ----------------------------
def fetch_posts():
    """Try API first. If it fails, open Chrome visibly."""
    try:
        response = requests.get(POSTS_API)
        response.raise_for_status()
        return response.json()[:MAX_POSTS]
    except:
        print("API unavailable, opening Chrome to get posts.")
        return fetch_posts_from_chrome()

def fetch_posts_from_chrome():
    pyautogui.press('win')
    time.sleep(0.5)
    pyautogui.write('chrome')
    pyautogui.press('enter')
    time.sleep(2)

    # Navigate to URL
    pyautogui.hotkey('ctrl', 'l')
    pyautogui.write(POSTS_API)
    pyautogui.press('enter')
    time.sleep(3)

    # Select all and copy
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.5)

    
    data = pyperclip.paste()
    posts = json.loads(data)
    print("Posts copied from Chrome successfully.")
    return posts[:MAX_POSTS]

# ----------------------------
# Main Automation
# ----------------------------
def main():
    posts = fetch_posts()

    for post in posts:
        for attempt in range(RETRY_ATTEMPTS):
            screenshot_path = take_screenshot()
            center = detect_icon(ICON_PATH, screenshot_path)
            if center:
                click_icon(center)
                if wait_for_notepad():
                    time.sleep(0.5)
                    type_and_save_post(post)
                    break
                else:
                    print("Notepad did not open, retrying...")
            else:
                print(f"Attempt {attempt+1}: Icon not found, retrying in {RETRY_DELAY}s...")
                time.sleep(RETRY_DELAY)
        else:
            print("Failed to find Notepad icon after retries. Skipping this post.")

if __name__ == "__main__":
    main()    
