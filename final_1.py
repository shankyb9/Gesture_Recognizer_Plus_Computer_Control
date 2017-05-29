import cv2
import numpy as np
import math
import cv2.cv as cv
from datetime import datetime
import time
from win32com.client import constants, Dispatch
import os
import webbrowser as wb
import process_check as pc
import PIL.Image
import serial
from Tkinter import *

WIDTHH=500
HEIGHTT=400

speaker = Dispatch("SAPI.SpVoice")

cap = cv2.VideoCapture(0)

class MotionDetectorInstantaneous():


    def onChange(self, val):  # callback when the user change the detection threshold
        self.threshold = val

    def __init__(self, threshold=70,  showWindows=True):

        self.writer = None
        self.font = None

        self.show = showWindows  # Either or not show the 2 windows
        self.frame = None

        self.capture = cv.CaptureFromCAM(0)
        self.frame = cv.QueryFrame(self.capture)  # Take a frame to init recorder

        self.frame=self.frame[1:100,540:640]
        self.frame1gray = cv.CreateMat(self.frame.height, self.frame.width, cv.CV_8U)  # Gray frame at t-1
        cv.CvtColor(self.frame, self.frame1gray, cv.CV_RGB2GRAY)

        # Will hold the thresholded result
        self.res = cv.CreateMat(self.frame.height, self.frame.width, cv.CV_8U)

        self.frame2gray = cv.CreateMat(self.frame.height, self.frame.width, cv.CV_8U)  # Gray frame at t

        self.width = self.frame.width
        self.height = self.frame.height
        self.nb_pixels = self.width * self.height
        self.threshold = threshold

        self.trigger_time = 0  # Hold timestamp of the last detection



        codec = cv.CV_FOURCC('M', 'J', 'P', 'G')  # ('W', 'M', 'V', '2')
        self.writer = cv.CreateVideoWriter(datetime.now().strftime("%b-%d_%H_%M_%S") + ".wmv", codec, 5,
                                           cv.GetSize(self.frame), 1)
        # FPS set to 5 because it seems to be the fps of my cam but should be ajusted to your needs
        self.font = cv.InitFont(cv.CV_FONT_HERSHEY_SIMPLEX, 1, 1, 0, 2, 8)  # Creates a font


    def run(self):
        started = time.time()
        menu=np.zeros((500,600,3), np.uint8)
        menu[5:492,7 :592] = (0, 100, 220)
        cv2.putText(menu, "MENU", (195, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, (255,255,255),2)
        cv2.putText(menu, "_____", (190, 60), cv2.FONT_HERSHEY_SIMPLEX, 2, (70, 70, 70), 2)
        cv2.putText(menu, "Button 0 : Number", (10, 100), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 0), 2)
        cv2.putText(menu, " 1: One  , 2: Two  , 3: Three", (10, 120), cv2.FONT_ITALIC, 0.5, (255, 255, 255), 1)
        cv2.putText(menu, "      4: Four  , 5: Five", (10, 140), cv2.FONT_ITALIC, 0.5, (255, 255, 255), 1)

        cv2.putText(menu, "Button 1 : OS Commands", (10, 170), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 0), 2)
        cv2.putText(menu, " 1: Play Music   , 2: Open Web  , 3: Kill Processes", (10, 190), cv2.FONT_ITALIC, 0.5, (255, 255, 255), 1)
        cv2.putText(menu, "    4: Open Text Files  , 5: Open Image Files", (10, 210), cv2.FONT_ITALIC, 0.5, (255, 255, 255), 1)

        cv2.putText(menu, "Button 2 : Sentences", (10, 240), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 0), 2)
        cv2.putText(menu, " 1: Hey  , 2: How r u ?   , 3: I am Fine !", (10, 260), cv2.FONT_ITALIC, 0.5, (255, 255, 255), 1)
        cv2.putText(menu, "          4: See u  , 5: Bye", (10, 280), cv2.FONT_ITALIC, 0.5, (255, 255, 255), 1)

        cv2.putText(menu, "Button 3 : Animations", (10, 310), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 0), 2)
        cv2.putText(menu, " 1: Horizontal Motion  , 2: Vertical Motion  , 3: Diagonal Motion", (10, 330), cv2.FONT_ITALIC, 0.5, (255, 255, 255), 1)
        cv2.putText(menu, "         4: Fast Motion  , 5: Slow Motion", (10, 350), cv2.FONT_ITALIC, 0.5, (255, 255, 255), 1)

        #cv2.putText(menu, "Button 4 : Arduino", (10, 380), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 0), 2)
        #cv2.putText(menu, " 1: LED On  , 2: LED Off  , 3: Obstacle Detection", (10, 400), cv2.FONT_ITALIC, 0.5, (255, 255, 255), 1)
        #cv2.putText(menu, "      4: Not Defined  , 5: Not Defined", (10, 420), cv2.FONT_ITALIC, 0.5, (255, 255, 255), 1)

        cv2.imshow('Menu', menu)
        cv2.moveWindow('Menu',0,0)
        c_one = 1
        c_two = 1
        c_three = 1
        c_four = 1
        c_zero = 1
        c_check = 7
        c_button=0
        c_reac_time=0
        path = 'E:\project_main\shanta'
        files = os.listdir(path)
        coun=0
        print "Gesture Recognizer Started :\n Button 0"
        process=('notepad.exe','WINWORD.EXE','IEXPLORE.exe','ois.exe','dllhost.exe','excel.exe','mpc-hc.exe','wmplayer.exe','wordpad.exe','calc.exe','powerpnt.exe','firefox.exe')
        fle = open('abc.txt', 'w')
        print >> fle, "----------------------------------OUTPUT OF THE GESTURES-----------------------------------"
        print >> fle,"---Instructions---"
        print >> fle, "Button 0 : Number"
        print >> fle, "Button 1 : OS Commands"
        print >> fle, "Button 2 : Sentences"
        print >> fle, "Button 3 : Animations"
        #print >> fle, "Button 4 : Arduino"
        print >> fle, "--------------------------------------------------------------------------------------------"
        print >> fle, "-------------------------------------------OUTPUT-------------------------------------------"
        print >> fle, "Button No. 0"
        str = ""

        while True:

            #######
            #######
            #######
            ret, img = cap.read()
            cv2.rectangle(img, (400, 400), (100, 100), (0, 255, 0), 0)
            crop_img = img[100:400, 100:400]
            grey = cv2.cvtColor(crop_img, cv2.COLOR_BGR2GRAY)
            value = (35, 35)
            blurred = cv2.GaussianBlur(grey, value, 0)
            _, thresh1 = cv2.threshold(blurred, 127, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
            #cv2.imshow('Thresholded', thresh1)

            contours, hierarchy = cv2.findContours(thresh1.copy(), cv2.RETR_TREE, cv2.CHAIN_APPROX_NONE)

            cnt = max(contours, key=lambda x: cv2.contourArea(x))

            x, y, w, h = cv2.boundingRect(cnt)
            cv2.rectangle(crop_img, (x, y), (x + w, y + h), (0, 0, 255), 0)
            hull = cv2.convexHull(cnt)
            #drawing = np.zeros(crop_img.shape, np.uint8)
            #cv2.drawContours(drawing, [cnt], 0, (0, 255, 0), 0)
            #cv2.drawContours(drawing, [hull], 0, (0, 0, 255), 0)
            hull = cv2.convexHull(cnt, returnPoints=False)

            count_defects = 0
            cv2.drawContours(thresh1, contours, -1, (0, 255, 0), 3)
            defects = cv2.convexityDefects(cnt, hull)
            for i in range(cv2.convexityDefects(cnt, hull).shape[0]):
                s, e, f, d = defects[i, 0]
                start = tuple(cnt[s][0])
                end = tuple(cnt[e][0])
                far = tuple(cnt[f][0])
                a = math.sqrt((end[0] - start[0]) ** 2 + (end[1] - start[1]) ** 2)
                b = math.sqrt((far[0] - start[0]) ** 2 + (far[1] - start[1]) ** 2)
                c = math.sqrt((end[0] - far[0]) ** 2 + (end[1] - far[1]) ** 2)
                angle = math.acos((b ** 2 + c ** 2 - a ** 2) / (2 * b * c)) * 57
                if angle <= 90:
                    count_defects += 1
                    cv2.circle(crop_img, far, 1, [0, 0, 255], -1)
                # dist = cv2.pointPolygonTest(cnt,far,True)
                cv2.line(crop_img, start, end, [0, 255, 0], 2)
                # cv2.circle(crop_img,far,5,[0,0,255],-1)
 
            c_reac_time=c_reac_time+1 
            if count_defects == 1:

                cv2.putText(img, str, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, 2)
                c_one = c_one + 1
                c_two = c_three = c_four = c_zero = 1
                if (c_one == c_check):

                    if(c_button==0):
                        str = "two"
                        cv2.putText(img, str, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, 2)
                    elif(c_button==1):
                        wb.open("www.rknec.edu")
                        str="Opened web"
                    elif(c_button==2):
                        str="How r u ?"
                    elif (c_button == 3):
                        tk = Tk()
                        xspeed=0
                        yspeed=3
                        canvas = Canvas(tk, width=WIDTHH, height=HEIGHTT)
                        tk.title("BALL ANIMATION")
                        canvas.pack()
                        coun = 1
                        ball = canvas.create_oval(10, 10, 60, 60, fill="blue")
                        while True:
                            try:
                                coun = coun + 1
                                if (coun == 500):
                                    tk.destroy()
                                    break
                                canvas.move(ball, xspeed, yspeed)
                                pos = canvas.coords(ball)
                                if pos[3] >= HEIGHTT or pos[1] <= 0:
                                    yspeed = -yspeed
                                # if pos[2] >= WIDTHH or pos[0] <= 0:
                                #     xspeed = -xspeed
                                tk.update()
                                time.sleep(0.01)

                            except Exception as e:
                                break

                        str = "Played Vertical Animation"
                        coun=0
                    cv2.putText(img, str, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, 2)
                    speaker.Speak(str)
                    print >> fle, str


            elif count_defects == 2:

                cv2.putText(img, str, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, 2)
                c_two = c_two + 1
                c_three = c_one = c_four = c_zero = 1
                if (c_two == c_check):

                    if (c_button == 0):
                        str = "Three"
                        cv2.putText(img, str, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 1, 2)
                    elif(c_button==1):
                        str = "Killed No Processes"
                        for i in process:
                            if (pc.processExists(i)):
                                 os.system("taskkill /f /im " + i)
                                 str = "Killed Process : " + i
                        cv2.putText(img, str, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 1, 1)
                    elif(c_button==2):
                        str="i m fine"

                    elif(c_button==3):
                        tk = Tk()
                        xspeed=3
                        yspeed=3
                        canvas = Canvas(tk, width=WIDTHH, height=HEIGHTT)
                        tk.title("BALL ANIMATION")
                        canvas.pack()
                        coun = 1
                        ball = canvas.create_oval(10, 10, 60, 60, fill="blue")
                        while True:
                            try:
                                coun=coun+1
                                if(coun==500):
                                    tk.destroy()
                                    break
                                canvas.move(ball, xspeed, yspeed)
                                pos = canvas.coords(ball)
                                if pos[3] >= HEIGHTT or pos[1] <= 0:
                                    yspeed = -yspeed
                                if pos[2] >= WIDTHH or pos[0] <= 0:
                                    xspeed = -xspeed
                                tk.update()
                                time.sleep(0.01)

                            except Exception as e:
                                break

                        str="Played Diagonal Animation"
                        coun=0
                    cv2.putText(img, str, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, 1)
                    speaker.Speak(str)
                    print >> fle, str

            elif count_defects == 3:

                cv2.putText(img, str, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, 2)
                c_three = c_three + 1
                c_two = c_one = c_four = c_zero = 1
                if (c_three == c_check):

                    if (c_button == 0):
                        str = "Four"
                        cv2.putText(img, str, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, 2)
                    elif(c_button==1):
                        files_txt = [i for i in files if i.endswith('.txt')]
                        for i in files_txt:
                            os.startfile( path + "\\" + i)
                        str="Opened text files"
                    elif(c_button==2):
                        str="see u"
                    elif (c_button == 3):
                        tk = Tk()
                        yspeed=7
                        xspeed=7
                        canvas = Canvas(tk, width=WIDTHH, height=HEIGHTT)
                        tk.title("BALL ANIMATION")
                        canvas.pack()
                        coun = 1
                        ball = canvas.create_oval(10, 10, 60, 60, fill="blue")
                        while True:
                            try:
                                coun = coun + 1
                                if (coun == 500):
                                    tk.destroy()
                                    break
                                canvas.move(ball, xspeed, yspeed)
                                pos = canvas.coords(ball)
                                if pos[3] >= HEIGHTT or pos[1] <= 0:
                                     yspeed = -yspeed
                                if pos[2] >= WIDTHH or pos[0] <= 0:
                                    xspeed = -xspeed
                                tk.update()
                                time.sleep(0.01)

                            except Exception as e:
                                break

                        str = "Played Fast Animation"
                        coun=0

                    cv2.putText(img, str, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, 2)
                    speaker.Speak(str)
                    print >> fle, str

            elif count_defects == 4:

                cv2.putText(img, str, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, 2)
                c_four = c_four + 1
                c_two = c_three = c_one = c_zero = 1
                if (c_four == c_check):

                    if (c_button == 0):
                        str = "Five"
                        cv2.putText(img, str, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, 2)
                    elif (c_button == 1):
                        for i in files:
                            if i.endswith('.JPG'):
                                imgs = PIL.Image.open(path + "\\" + i)
                                imgs.show()

                        str = "Opened image files"

                    elif(c_button==2):
                        str="bye"
                    elif (c_button == 3):
                        tk = Tk()
                        yspeed=1
                        xspeed=1
                        canvas = Canvas(tk, width=WIDTHH, height=HEIGHTT)
                        tk.title("BALL ANIMATION")
                        canvas.pack()
                        coun = 1
                        ball = canvas.create_oval(10, 10, 60, 60, fill="blue")
                        while True:
                            try:
                                coun = coun + 1
                                if (coun == 500):
                                    tk.destroy()
                                    break
                                canvas.move(ball, xspeed, yspeed)
                                pos = canvas.coords(ball)
                                if pos[3] >= HEIGHTT or pos[1] <= 0:
                                     yspeed = -yspeed
                                if pos[2] >= WIDTHH or pos[0] <= 0:
                                    xspeed = -xspeed
                                tk.update()
                                time.sleep(0.01)

                            except Exception as e:
                                break

                        str = "Played Slow Animation"
                        coun=0

                    cv2.putText(img, str, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, 2)
                    speaker.Speak(str)
                    print >> fle, str

            elif count_defects == 0:

                cv2.putText(img, str, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, 2)
                c_zero = c_zero + 1
                c_two = c_three = c_four = c_one = 1
                if (c_zero == c_check):

                    if (c_button == 0):
                        str = "one"
                        cv2.putText(img, str, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, 2)
                    elif(c_button==1):
                        os.system('start ratta.mp3')
                        str = "Playing music"
                    elif(c_button==2):
                        str="Hey"
                    elif (c_button == 3):
                        tk = Tk()
                        yspeed=0
                        xspeed=5
                        canvas = Canvas(tk, width=WIDTHH, height=HEIGHTT)
                        tk.title("BALL ANIMATION")
                        canvas.pack()
                        coun = 1
                        ball = canvas.create_oval(10, 10, 60, 60, fill="blue")
                        while True:
                            try:
                                coun = coun + 1
                                if (coun == 500):
                                    tk.destroy()
                                    break
                                canvas.move(ball, xspeed, yspeed)
                                pos = canvas.coords(ball)
                                # if pos[3] >= HEIGHTT or pos[1] <= 0:
                                #     yspeed = -yspeed
                                if pos[2] >= WIDTHH or pos[0] <= 0:
                                    xspeed = -xspeed
                                tk.update()
                                time.sleep(0.01)

                            except Exception as e:
                                break

                        str = "Played Horizontal Animation"
                        coun=0



                    cv2.putText(img, str, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 2, 1)
                    speaker.Speak(str)
                    print >>fle, str


            img[1:100,540:640]=(0,0,255)
            cv2.imshow('Gesture', img)
            cv2.moveWindow('Gesture',650,0)
            #all_img = np.hstack((drawing, crop_img))
            #cv2.imshow('Contours', all_img)
            #####
            #####
            #####



            curframe1 = cv.QueryFrame(self.capture)
            curframe=curframe1[1:100,540:640]


            instant = time.time()  # Get timestamp o the frame

            self.processImage(curframe)  # Process the image

    #        if not self.isRecording:
            if self.somethingHasMoved():
                    self.trigger_time = instant  # Update the trigger_time
                    if instant > started + 5:  # Wait 5 second after the webcam start for luminosity adjusting etc..
                        if c_reac_time>=10:
                            c_button=np.mod(c_button+1,4)
                            c_reac_time=0
                            print datetime.now().strftime("%b %d, %H:%M:%S"), "Button No. " ,c_button
                            print >> fle, datetime.now().strftime("%b %d, %H:%M:%S"), "Button No. " ,c_button


            cv.Copy(self.frame2gray, self.frame1gray)
            c = cv.WaitKey(1) % 0x100
            if c == 27 :
                break

    def processImage(self, frame):
        cv.CvtColor(frame, self.frame2gray, cv.CV_RGB2GRAY)

        # Absdiff to get the difference between to the frames
        cv.AbsDiff(self.frame1gray, self.frame2gray, self.res)

        # Remove the noise and do the threshold
        cv.Smooth(self.res, self.res, cv.CV_BLUR, 5, 5)
        cv.MorphologyEx(self.res, self.res, None, None, cv.CV_MOP_OPEN)
        cv.MorphologyEx(self.res, self.res, None, None, cv.CV_MOP_CLOSE)
        cv.Threshold(self.res, self.res, 10, 255, cv.CV_THRESH_BINARY_INV)

    def somethingHasMoved(self):
        nb = 0  # Will hold the number of black pixels

        for x in range(self.height):  # Iterate the hole image
            for y in range(self.width):
                if self.res[x, y] == 0.0:  # If the pixel is black keep it
                    nb += 1
        avg = (nb * 100.0) / self.nb_pixels  # Calculate the average of black pixel in the image

        if avg > self.threshold:  # If over the ceiling trigger the alarm
            return True
        else:
            return False


if __name__ == "__main__":
    detect = MotionDetectorInstantaneous()
    detect.run()


del speaker
