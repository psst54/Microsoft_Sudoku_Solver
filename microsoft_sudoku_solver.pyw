import win32gui
import win32con
from PIL import ImageGrab, Image
import cv2
import numpy as np
import sys
import os
import pyautogui

def getImage():
    hwnd = win32gui.FindWindow(None, "Microsoft Sudoku")
    if not hwnd:
        print("MIcrosoft Sudoku is not running")
        return False
    win32gui.SetForegroundWindow(hwnd)

    global win_x1; global win_x2; global win_y1; global win_y2
    (win_x1, win_y1, win_x2, win_y2) = win32gui.GetClientRect(hwnd)
    win_x1, win_y1 = win32gui.ClientToScreen(hwnd, (win_x1, win_y1))
    win_x2, win_y2 = win32gui.ClientToScreen(hwnd, (win_x2, win_y2))

    img = ImageGrab.grab((win_x1, win_y1, win_x2, win_y2))
    img.save('./image.png')

    return True
    

def loadTemplate():
    for i in range(10):
        template_img_number.append((cv2.imread((f'./template/{i}.png'), cv2.IMREAD_GRAYSCALE), i))
    for i in range(10):
        template_img_shapes.append((cv2.imread((f'./template/c{i}.png'), cv2.IMREAD_GRAYSCALE), i))
    return

    
def readNumber(img):
    img_h, img_w = img.shape[:2]

    max_sad = 100000
    number = 0

    for i in range(10):

        template_h, template_w = template_img_number[i][0].shape[:2]

        for img_x in range(img_w - template_w):
            for img_y in range(img_h - template_h):
                SAD = 0
                
                img_pixel = img[img_y:img_y + template_h , img_x:img_x + template_w].ravel()
                template_pixel = template_img_number[i][0].ravel()

                SAD = np.abs(np.subtract(img_pixel, template_pixel, dtype = np.float))
                SAD = SAD.sum()

                if max_sad > SAD:
                    max_sad = SAD
                    number = template_img_number[i][1]

    return number


def readShapes(img):
    img_h, img_w = img.shape[:2]

    max_sad = 100000
    number = 0

    for i in range(10):

        template_h, template_w = template_img_shapes[i][0].shape[:2]

        for img_x in range(img_w - template_w):
            for img_y in range(img_h - template_h):
                SAD = 0
                
                img_pixel = img[img_y:img_y + template_h , img_x:img_x + template_w].ravel()
                template_pixel = template_img_shapes[i][0].ravel()

                SAD = np.abs(np.subtract(img_pixel, template_pixel, dtype = np.float))
                SAD = SAD.sum()

                if max_sad > SAD:
                    max_sad = SAD
                    number = template_img_shapes[i][1]
    return number


def imgToArray():
    img = cv2.imread('./image.png')

    img_gray = cv2.inRange(img, (0, 0, 0), (150, 200, 10))

    contours, hierarchy = cv2.findContours(img_gray, cv2.RETR_CCOMP, cv2.CHAIN_APPROX_SIMPLE)
    for i in range(len(contours) - 1, -1, -1):
        chk = True
        for j in range(8):
            if i - j - 1 != hierarchy[0][i - j][1]:
                chk = False
                break
        if chk:
            start_idx = i - 8
            end_idx = i
            break

    # find x, y of board
    global board_x
    global board_y
    board_x = set()
    board_y = set()
    for i in range(start_idx, end_idx + 1):
        x_idx = 987654321
        y_idx = 987654321
        for j in range(len(contours[i])):
            x_idx = min(x_idx, contours[i][j][0][0])
            y_idx = min(y_idx, contours[i][j][0][1])
        board_x.add(x_idx)
        board_y.add(y_idx)
    board_x = list(board_x)
    board_y = list(board_y)
    board_x.sort()
    board_y.sort()

    cell_size = int((board_x[1] - board_x[0]) / 3)

    for i in range(3):
        for j in range(1, 3):
            board_x[i*3+j:i*3+j] = [board_x[i*3+j-1] + cell_size]
            board_y[i*3+j:i*3+j] = [board_y[i*3+j-1] + cell_size]

    # read board
    sudoku_board = [[] for _ in range(9)]
    zero_counter = 0
    img_gray = cv2.inRange(img, (0, 0, 0), (0, 0, 0))
    for i in range(9):
        for j in range(9):
            cell_img = img_gray[board_y[i]:board_y[i] + cell_size, board_x[j]:board_x[j] + cell_size].copy()
            cell_img = cv2.resize(cell_img, dsize=(30, 30), interpolation=cv2.INTER_AREA)
            num = readNumber(cell_img)
            sudoku_board[i].append(num)
            if not num:
                zero_counter += 1
            
    if zero_counter != 81:
        return sudoku_board
    
    sudoku_board = [[] for _ in range(9)]
    img_gray = cv2.inRange(img, (0, 20, 20), (255, 200, 255))
    for i in range(9):
        for j in range(9):
            cell_img = img_gray[board_y[i]:board_y[i] + cell_size, board_x[j]:board_x[j] + cell_size].copy()
            cell_img = cv2.resize(cell_img, dsize=(30, 30), interpolation=cv2.INTER_AREA)
            num = readShapes(cell_img)
            sudoku_board[i].append(num)
            
    return sudoku_board


def recur(n):
    global chk
    
    if (n == cnt):
        chk = True
        return

    x = idx[n][0]
    y = idx[n][1]
    for i in range(1, 10):
        if row[x][i] or col[y][i] or cell[(y // 3) * 3 + (x // 3)][i]:
            continue
        row[x][i] = col[y][i] = cell[(y // 3) * 3 + (x // 3)][i] = True
        sudoku_board[y][x] = i
        recur(n + 1)
        if chk:
            return
        row[x][i] = col[y][i] = cell[(y // 3) * 3 + (x // 3)][i] = False
    return


def click():
    img = cv2.imread('./image.png')
    os.remove('./image.png')
    img_gray = cv2.inRange(img, (0, 0, 0), (150, 200, 10))
    contours, hierarchy = cv2.findContours(img_gray, cv2.RETR_CCOMP, cv2.CHAIN_APPROX_SIMPLE)
    for i in range(len(contours)):
        chk = True
        for j in range(8):
            if i + j != hierarchy[0][i + j][1] + 1:
                chk = False
                break
        if chk:
            start_idx = i
            end_idx = i + 8
            break

    button_y = 987654321
    for i in range(start_idx, end_idx + 1):
        for j in range(len(contours[start_idx + i])):
            button_y = min(button_y, contours[i][j][0][1])
    margin = 20
  
    for i in range(1, 10):
        pyautogui.click(win_x1 + board_x[i - 1] + margin, win_y1 + button_y - margin * 2, button='left', clicks=1)
        for j in range(9):
            for k in range(9):
                if sudoku_board_original[j][k] == 0 and sudoku_board[j][k] == i:
                    pyautogui.click(win_x1 + board_x[k] + margin, win_y1 + board_y[j] + margin, button='left', clicks=1)

    return

if __name__ == "__main__":
    win_x1, win_x2, win_y1, win_y2 = (0, 0, 0, 0)
    template_img_number = []
    template_img_shapes = []
    board_x, board_y = (0, 0)
    row = [[False] * 10 for _ in range(9)]
    col = [[False] * 10 for _ in range(9)]
    cell = [[False] * 10 for _ in range(9)]
    idx = [[False] * 2 for _ in range(81)]
    cnt = 0
    chk = False

    if (not getImage()):
        sys.exit()

    loadTemplate()

    sudoku_board = imgToArray()
    sudoku_board_original = [[[False] for _ in range(9)] for _ in range(9)]
    for i in range(9):
        for j in range(9):
            sudoku_board_original[i][j] = sudoku_board[i][j]
            if (sudoku_board[i][j]):
                col[i][sudoku_board[i][j]] = True
                row[j][sudoku_board[i][j]] = True
                cell[(i // 3) * 3 + (j // 3)][sudoku_board[i][j]] = True
            else:
                idx[cnt][0] = j
                idx[cnt][1] = i
                cnt += 1
        
    recur(0)

    click()
