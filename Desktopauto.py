import os
import time
import json
import requests
import pyautogui
import pygetwindow as gw
import pyperclip
import cv2
import numpy as np
from botcity.core import DesktopBot
from PIL import Image
from pywinauto import keyboard
from pywinauto.application import Application

# ------------------ CONSTANTS ------------------

ICON_PATH = "notepad_icon.png"  # Path to the Notepad icon image
OUTPUT_DIR = r"C:\Users\amrk6\Desktop\tjm-project"
ANNOTATED_DIR = os.path.join(OUTPUT_DIR, "annotated")
POSTS_API = "https://jsonplaceholder.typicode.com/posts"
MAX_POSTS = 10
RETRY_ATTEMPTS = 3
RETRY_DELAY = 1  # seconds

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(ANNOTATED_DIR, exist_ok=True)

# ------------------ UTILITY FUNCTIONS ------------------

def close_unexpected_popups(main_window_title):
    """Closes all windows except the main one."""
    windows = gw.getAllTitles()
    for w in windows:
        if w and main_window_title not in w:
            try:
                win = gw.getWindowsWithTitle(w)[0]
                win.activate()
                time.sleep(0.3)
                pyautogui.press('esc')
                time.sleep(0.2)
                pyautogui.hotkey('alt', 'f4')
                time.sleep(0.2)
            except Exception as e:
                print(f"Could not close window '{w}': {e}")

def fetch_posts():
    """Fetch posts from API, fallback to Chrome if API unavailable."""
    try:
        response = requests.get(POSTS_API)
        response.raise_for_status()
        return response.json()[:MAX_POSTS]
    except:
        print("API unavailable, opening Chrome to fetch posts.")
        return fetch_posts_from_chrome()

def fetch_posts_from_chrome():
    """Fetch posts by opening Chrome and copying the JSON."""
    pyautogui.hotkey('win')
    time.sleep(0.5)
    pyautogui.write('chrome')
    pyautogui.press('enter')
    time.sleep(2)

    pyautogui.hotkey('ctrl', 'l')
    pyautogui.write(POSTS_API)
    pyautogui.press('enter')
    time.sleep(3)

    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.5)
    
    data = pyperclip.paste()
    posts = json.loads(data)
    print("Posts copied from Chrome successfully.")
    return posts[:MAX_POSTS]

# ------------------ NOTEPAD FUNCTIONS ------------------

def open_notepad(bot: DesktopBot, screenshot_index=0, scales=[0.5, 0.75, 1.0, 1.25, 1.5, 2.0], threshold=0.5):
    """Open Notepad by detecting its icon robustly on the desktop."""
    
    # Minimize all windows
    bot.type_keys(["win", "d"])
    time.sleep(1)

    # Take a desktop screenshot
    screenshot_path = os.path.join(ANNOTATED_DIR, f"annotated_screenshot_{screenshot_index}.png")
    bot.screenshot(screenshot_path)
    desktop_color = cv2.imread(screenshot_path)
    desktop_gray = cv2.cvtColor(desktop_color, cv2.COLOR_BGR2GRAY)

    # Load target icon
    target_img = cv2.imread(ICON_PATH, cv2.IMREAD_GRAYSCALE)
    if target_img is None:
        print("Icon image not found.")
        return False

    # --- Multi-scale template matching ---
    target_edges = cv2.Canny(target_img, 50, 150)
    desktop_edges = cv2.Canny(desktop_gray, 50, 150)
    for scale in scales:
        resized_target = cv2.resize(target_edges, (0, 0), fx=scale, fy=scale)
        res = cv2.matchTemplate(desktop_edges, resized_target, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(res)
        if max_val >= threshold:
            x, y = max_loc
            h, w = resized_target.shape
            center_x = x + w // 2
            center_y = y + h // 2
            bot.mouse_move(center_x, center_y)
            pyautogui.doubleClick()
            print(f"Notepad icon found via template matching at ({center_x}, {center_y})")
            return True

    # --- ORB feature matching as fallback ---
    orb = cv2.ORB_create()
    kp1, des1 = orb.detectAndCompute(target_img, None)
    kp2, des2 = orb.detectAndCompute(desktop_gray, None)
    if des1 is not None and des2 is not None:
        bf = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
        matches = bf.match(des1, des2)
        matches = sorted(matches, key=lambda x: x.distance)
        if len(matches) > 10:  # enough good matches
            src_pts = np.float32([kp1[m.queryIdx].pt for m in matches]).reshape(-1,1,2)
            dst_pts = np.float32([kp2[m.trainIdx].pt for m in matches]).reshape(-1,1,2)
            avg_x = int(np.mean(dst_pts[:,0,0]))
            avg_y = int(np.mean(dst_pts[:,0,1]))
            bot.mouse_move(avg_x, avg_y)
            pyautogui.doubleClick()
            print(f"Notepad icon found via ORB feature matching at ({avg_x}, {avg_y})")
            return True

    print("Notepad icon not detected on desktop.")
    return False

def fallback_open_notepad_via_search(main_window_title="Untitled - Notepad"):
    """Open Notepad via Windows search and close popups."""
    print("Fallback: Opening Notepad via Windows search...")
    pyautogui.hotkey('win')
    time.sleep(0.5)
    pyautogui.write("Notepad")
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(1.5)
    close_unexpected_popups(main_window_title)

def wait_for_notepad(timeout=10):
    """Wait for Notepad window to become active."""
    start_time = time.time()
    while time.time() - start_time < timeout:
        windows = gw.getWindowsWithTitle("Untitled - Notepad")
        if windows:
            try:
                windows[0].activate()
                time.sleep(0.5)
            except:
                pass
            if windows[0].isActive:
                return windows[0]
        time.sleep(0.5)
    return None

def type_and_save_post(post):
    """Type post content into Notepad and save as text file."""
    content = f"Title: {post['title']}\n\n{post['body']}\n\n"

    for line in content.splitlines():
        pyautogui.write(line, interval=0.03)
        pyautogui.press('enter')
    time.sleep(0.5)

    base_name = f"post_{post['id']}"
    full_path = os.path.join(OUTPUT_DIR, f"{base_name}.txt")
    counter = 1
    while os.path.exists(full_path):
        full_path = os.path.join(OUTPUT_DIR, f"{base_name}_{counter}.txt")
        counter += 1

    pyautogui.hotkey('ctrl', 's')
    time.sleep(0.5)
    pyautogui.write(full_path)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'w')

# ------------------ MAIN EXECUTION ------------------

if __name__ == "__main__":
    bot = DesktopBot()
    main_window_title = "Untitled - Notepad"

    posts = fetch_posts()
    if not posts:
        print("No posts available.")
        exit(1)

    for idx, post in enumerate(posts):
        success = False
        for attempt in range(RETRY_ATTEMPTS):
            if open_notepad(bot, screenshot_index=idx):
                notepad_window = wait_for_notepad()
                if notepad_window:
                    success = True
                    break
            print(f"Attempt {attempt + 1} failed to open Notepad. Retrying in {RETRY_DELAY} sec...")
            time.sleep(RETRY_DELAY)

        if not success:
            fallback_open_notepad_via_search(main_window_title)
            notepad_window = wait_for_notepad()
            if notepad_window:
                success = True

        if not success:
            print(f"Failed to open Notepad for post {post['id']}. Skipping...")
            continue

        type_and_save_post(post)
        print(f"Post {post['id']} typed and saved successfully.")
        time.sleep(1)

    print("All posts processed!")
