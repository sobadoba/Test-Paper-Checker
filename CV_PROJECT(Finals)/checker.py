import os
import subprocess
import sys

import matplotlib.pyplot as plt
import pytesseract
import xlsxwriter
import cv2 as cv
import numpy as np


def roundedContours(contours):
    circCont = []
    for i in contours:
        area = cv.contourArea(i)
        # check area to adjust size
        # print(area)
        if 150 < area < 300:
            approx = cv.approxPolyDP(i, .03 * cv.arcLength(i, True), True)
            # print(len(approx))
            if len(approx) >= 6:
                circCont.append(i)
    # print(len(circCont))
    return circCont


def shadedCircles(image):
    img_gray = cv.cvtColor(image, cv.COLOR_BGR2GRAY)
    contours, hierarchy = cv.findContours(img_gray, cv.RETR_LIST, cv.CHAIN_APPROX_NONE)
    # test
    image_contours = image.copy()
    cv.drawContours(image_contours, contours, -1, (0, 255, 0), 2)
    # cv.imshow("image_contours", image_contours)

    # get all rounded contours
    circles = roundedContours(contours)

    return circles


def scoring(image, contours):
    num_items_right = 50  # change value nalang
    for i in contours:
        x, y, w, h = cv.boundingRect(i)
        # print(x, y, w, h) # test

        # get roi of circle
        roi = image[y:y + h, x:x + w]

        num_white_pix = np.sum(roi == 255)
        # print(num_white_pix) # test to see value

        if num_white_pix > 300:  # change nalang if ever
            cv.rectangle(image, (x + w + 2, y + h), (x - 2, y), (0, 0, 255), 1)
            num_items_right = num_items_right - 1
        else:
            cv.rectangle(image, (x + w + 2, y + h), (x - 2, y), (0, 255, 0), 1)

    return num_items_right


def checking(answer_key, folder):
    answer_key = cv.imread(answer_key)
    answer_key = cv.resize(answer_key, (790, 515))
    circles = shadedCircles(answer_key)

    j = 1
    path = "checked papers/"
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\tesseract.exe" # Modify the path where the tesseract.exe is found
    workbook = xlsxwriter.Workbook('Papers.xlsx')

    worksheet = workbook.add_worksheet("My sheet")

    worksheet.write(0, 0, 'Subject: ')
    worksheet.write(1, 0, 'Name')
    worksheet.write(1, 1, 'Grade and Section')
    worksheet.write(1, 2, 'Score')

    row = 2
    col = 0
    for img in os.listdir(folder):
        img = cv.imread(os.path.join(folder, img))
        test_sheet = img
        test_sheet = cv.resize(test_sheet, (790, 515))

        final = scoring(test_sheet, circles)

        text = pytesseract.image_to_string(img)
        per_line = text.split("\n")
        lines = per_line[0:5]
        lines.remove("")
        lines.remove("")
        lines = [lines]
        cv.putText(test_sheet, "SCORE: ", (int(500), int(70)), cv.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 0,), 2)
        label = cv.putText(test_sheet, str(final), (int(630), int(70)),
                           cv.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
        new = cv.circle(label, (int(650), int(60)), 30, (0, 0, 255), 2)

        cv.imwrite(os.path.join(path, "checked_testpaper" + str(j) + ".jpg"), new)
        plt.figure("Checked Test Paper " + str(j))
        plt.subplot(1, 2, 1)
        plt.imshow(answer_key)
        plt.subplot(1, 2, 2)
        plt.imshow(test_sheet)
        plt.show()
        j += 1

        for name, score, sc in lines:
            x = name.split(": ")
            y = score.split(": ")
            worksheet.write(row, col, x[1])
            worksheet.write(row, col + 1, y[1])
            worksheet.write(row, col + 2, final)
            row += 1
    workbook.close()
    done = cv.imread("assets/done.png")
    done = cv.resize(done, (400, 300))
    cv.imshow("Done Checking", done)
    path = r'C:\Users\shish\PycharmProjects\CV_PROJECT(Finals)\checked papers' # Modify the path where the checked test papers is stored or found in a folder
    sys.path.append(path)
    subprocess.Popen('explorer '+path)