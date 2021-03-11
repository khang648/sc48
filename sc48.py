########################################################## Import Module ################################################################
from tkinter import *
from tkinter import messagebox
import time
from time import sleep, gmtime, strftime
#from picamera import PiCamera
import cv2 
import numpy as np
from tkinter import filedialog as fd
from PIL import ImageTk, Image
#import serial
from functools import partial
import math
from fractions import Fraction
from threading import *
import os
from tkinter import ttk
import awesometkinter as atk
import tkinter.font as font
# from gpiozero import LED
# tx = LED(14)
# rx = LED(15)
# rx.off()
# tx.off()
# sleep(1)

######################################################### Khởi tạo Main Window ##############################################################
root = Tk()
root.title("SpotCheck Application")
root.geometry('850x450')
root.configure(background = "white")
s = ttk.Style()
s.theme_use('default')

########################################################### Sắp xếp contours ###############################################################
def sorting_y(contour):
    rect_y = cv2.boundingRect(contour)
    return rect_y[1]
def sorting_x(contour):
    rect_x = cv2.boundingRect(contour)
    return rect_x[0]
def sorting_xy(contour):
    rect_xy = cv2.boundingRect(contour)
    return math.sqrt(math.pow(rect_xy[0],2) + math.pow(rect_xy[1],2))

############################################################# Xử lý ảnh ###################################################################
def process_image(image_name):
    image = cv2.imread(image_name)
    blur_img = cv2.GaussianBlur(image.copy(), (33,33), 0)            
    gray = cv2.cvtColor(blur_img, cv2.COLOR_BGR2GRAY)
    gray_img = gray.copy()
    thresh, binary_img = cv2.threshold(gray_img, 40, maxval=255, type=cv2.THRESH_BINARY) 
    contours, hierarchy = cv2.findContours(binary_img, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    print("Number of contours: " + str(len(contours)))

    contours.sort(key=lambda data:sorting_xy(data))

    contour_img = np.zeros_like(gray_img)
    bourect0 = cv2.boundingRect(contours[0])
    bourect47 = cv2.boundingRect(contours[len(contours)-1])
    start_point = (bourect0[0]-15, bourect0[1]-15)  
    end_point = (bourect47[0]+bourect47[2]+15, bourect47[1]+bourect47[3]+15)
    contour_img = cv2.rectangle(contour_img, start_point, end_point, (255,255,255), -1)
    rect_w = (bourect47[0]+bourect47[2]+15) - (bourect0[0]-15)
    rect_h = (bourect47[1]+bourect47[3]+15) - (bourect0[1]-15)
    cell_w = round(rect_w/6)
    cell_h = round(rect_h/8)
    for i in range(1,6):
        contour_img = cv2.line(contour_img, (start_point[0]+i*cell_w,start_point[1]), (start_point[0]+i*cell_w,end_point[1]),(0,0,0), 3)
    for i in range(1,8):
        contour_img = cv2.line(contour_img, (start_point[0],start_point[1]+i*cell_h), (end_point[0],start_point[1]+i*cell_h),(0,0,0), 3)

    #gray1_img = cv2.cvtColor(contour_img, cv2.COLOR_BGR2GRAY)
    thresh1 , binary1_img = cv2.threshold(contour_img, 200, maxval=255, type=cv2.THRESH_BINARY)
    contours1, hierarchy1 = cv2.findContours(binary1_img, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)

    contours1.sort(key=lambda data:sorting_y(data))
    contours1_h1 = contours1[0:6]
    contours1_h2 = contours1[6:12]
    contours1_h3 = contours1[12:18]
    contours1_h4 = contours1[18:24]
    contours1_h5 = contours1[24:30]
    contours1_h6 = contours1[30:36]
    contours1_h7 = contours1[36:42]
    contours1_h8 = contours1[42:48]
    contours1_h1.sort(key=lambda data:sorting_x(data))
    contours1_h2.sort(key=lambda data:sorting_x(data))
    contours1_h3.sort(key=lambda data:sorting_x(data))
    contours1_h4.sort(key=lambda data:sorting_x(data))
    contours1_h5.sort(key=lambda data:sorting_x(data))
    contours1_h6.sort(key=lambda data:sorting_x(data))
    contours1_h7.sort(key=lambda data:sorting_x(data))
    contours1_h8.sort(key=lambda data:sorting_x(data))

    sorted_contours1 = contours1_h1 + contours1_h2 + contours1_h3 + contours1_h4 + contours1_h5 + contours1_h6 + contours1_h7 + contours1_h8
                 
    list_intensities = []
    sum_intensities = []
    result_list = list(range(48))
    area = list(range(48))

    for i in range(len(sorted_contours1)):
        cimg = np.zeros_like(gray_img)
        cv2.drawContours(cimg, sorted_contours1, i, color = 255, thickness = -1)
        pts = np.where(cimg == 255)
        list_intensities.append(gray_img[pts[0], pts[1]])
        sum_intensities.append(sum(list_intensities[i]))
        area[i]= cv2.contourArea(sorted_contours1[i])
        result_list[i] = round((sum_intensities[i])*50/69791)

#     blur1_img = cv2.GaussianBlur(image.copy(), (33,33), 0) 
#     hsv_img = cv2.cvtColor(blur1_img, cv2.COLOR_BGR2HSV)
#     list_hsvvalue = []
#     list_index = list(range(48))
#     for i in range(len(sorted_contours1)):
#         list_index[i] = []
#         cimg = np.zeros_like(gray_img)
#         cv2.drawContours(cimg, sorted_contours1, i, color = 255, thickness = -1)
#         pts = np.where(cimg == 255)
#         list_hsvvalue.append(hsv_img[pts[0], pts[1]])
#         for j in range(len(list_hsvvalue[i])):
#             list_index[i].append(list_hsvvalue[i][j][2])
#         list_intensities.append(sum(list_index[i]))
#         area[i]= cv2.contourArea(sorted_contours1[i])
#         result_list[i] = round((list_intensities[i])*45/59768)

    for i in range(len(sorted_contours1)):
        if ((i!=0) and ((i+1)%6==0)):
            print('%d' %(result_list[i]))
        else:
            print('%d' % (result_list[i]), end = ' | ')

    blurori_img = cv2.GaussianBlur(image.copy(), (25,25), 0)
    for i in range(len(sorted_contours1)):
        if(result_list[i] < 20):
            cv2.drawContours(blurori_img, sorted_contours1, i, (255,255,0), thickness = 2)
        else:
            cv2.drawContours(blurori_img, sorted_contours1, i, (0,0,255), thickness = 2)
#     for i in range(len(contours)):
#         cv2.drawContours(blurori_img, contours, i, (255,255,255), thickness = 1)

    return (result_list, blurori_img, start_point, end_point)

########################################################## Giao diện chính ################################################################
def mainscreen():
	buttonFont = font.Font(family='Helvetica', size=10, weight='bold')	
	global mainscreen_labelframe
	mainscreen_labelframe = LabelFrame(root, bg='white', width=850, height=417)
	mainscreen_labelframe.place(x=0,y=0)

	sidebar_labelframe = LabelFrame(mainscreen_labelframe, bg='DodgerBlue2', width=170, height=414)
	sidebar_labelframe.place(x=0,y=0)
	
	def covid19_click():
		covid19_canvas['bg'] = 'white'
		tb_canvas['bg'] = 'DodgerBlue2'
		viewresult_canvas['bg'] = 'DodgerBlue2'

		global covid19clicked
		covid19clicked = 2
		global tbclicked
		tbclicked = 0
		global viewresultclicked
		viewresultclicked = 0

		if(covid19_canvas['bg']=='white'):
			covid19mc_labelframe = LabelFrame(mainscreen_labelframe, bg='white', width=675, height=414)
			covid19mc_labelframe.place(x=172,y=0)

			def newprogram_click():
				mainscreen_labelframe.place_forget()
				settemp()
			newprogram_button = Button(covid19mc_labelframe, bg="lavender", text="NEW PROGRAM", height=3, width=15, command=newprogram_click)
			newprogram_button.place(x=272,y=250)

	def tb_click():
		covid19_canvas['bg'] = 'DodgerBlue2'
		tb_canvas['bg'] = 'white'
		viewresult_canvas['bg'] = 'DodgerBlue2'

		global covid19clicked
		covid19clicked = 0
		global tbclicked
		tbclicked = 1
		global viewresultclicked
		viewresultclicked = 0

		if(tb_canvas['bg']=='white'):
			tbmc_labelframe = LabelFrame(mainscreen_labelframe, bg='yellow', width=675, height=414)
			tbmc_labelframe.place(x=172,y=0)

	def viewresult_click():
		covid19_canvas['bg'] = 'DodgerBlue2'
		tb_canvas['bg'] = 'DodgerBlue2'
		viewresult_canvas['bg'] = 'white'

		global covid19clicked
		covid19clicked = 0
		global tbclicked
		tbclicked = 0
		global viewresultclicked
		viewresultclicked = 1

		if(viewresult_canvas['bg']=='white'):
			viewresultmc_labelframe = LabelFrame(mainscreen_labelframe, bg='red', width=675, height=414)
			viewresultmc_labelframe.place(x=172,y=0)
	
	covid19_button = Button(mainscreen_labelframe, bg="DodgerBlue2", text="COVID 19", fg='white', font=buttonFont, borderwidth=0, height=4, width=20, command=covid19_click)
	covid19_button.place(x=1,y=90)
	tb_button = Button(mainscreen_labelframe, bg="DodgerBlue2", text="TB", fg='white', font=buttonFont, borderwidth=0, height=4, width=20, command=tb_click)
	tb_button.place(x=1,y=168)
	viewresult_button = Button(mainscreen_labelframe, bg="DodgerBlue2", text="VIEW RESULT", fg='white', font=buttonFont, borderwidth=0, height=4, width=20, command=viewresult_click)
	viewresult_button.place(x=1,y=246)
	covid19_canvas = Canvas(mainscreen_labelframe, bg="DodgerBlue2", bd=0, highlightthickness=0, height=72, width=13)
	covid19_canvas.place(x=1,y=90)
	tb_canvas = Canvas(mainscreen_labelframe, bg="DodgerBlue2", bd=0, highlightthickness=0, height=72, width=13)
	tb_canvas.place(x=1,y=168)
	viewresult_canvas = Canvas(mainscreen_labelframe, bg="DodgerBlue2", bd=0, highlightthickness=0, height=72, width=13)
	viewresult_canvas.place(x=1,y=246)

###################################################### Giao diện set nhiệt độ ############################################################ 
def settemp():
	settemp_labelframe = LabelFrame(root, bg='white', width=850, height=417)
	settemp_labelframe.place(x=0,y=0)
	settemptop_labelframe = LabelFrame(settemp_labelframe, bg='white', width=846, height=288)
	settemptop_labelframe.place(x=0,y=45)
	#settempside_labelframe = LabelFrame(settemp_labelframe, bg='DodgerBlue2', width=396, height=413)
	#settempside_labelframe.place(x=450,y=0)
	keypad_labelframe = LabelFrame(settemptop_labelframe, bg='white', width=264, height=259)
	keypad_labelframe.place(x=520,y=12)
	settemp_label = Label(settemp_labelframe, bg='white', fg='black', text='SET TEMPERATURE', font=("Courier",17,'bold'), width=20, height=1 )
	settemp_label.place(x=290,y=7)

	def numpad_click(btn):
		text = "%s" % btn
		if (text!="Delete" and text!="Default"):
			if(entry_num==1):
				t1_entry.insert(END, text)
			if(entry_num==2):
				t2_entry.insert(END, text)
			if(entry_num==3):
				t3_entry.insert(END, text)
		if text == 'Delete':
			if(entry_num==1):
				t1_entry.delete(0, END)
			if(entry_num==2):
				t2_entry.delete(0, END)
			if(entry_num==3):
				t3_entry.delete(0, END)
		if text == 'Default':
			if(entry_num==1):
				t1_entry.delete(0, END)
				t1_entry.insert(END, t1)
			if(entry_num==2):
				t2_entry.delete(0, END)
				t2_entry.insert(END, t2)
			if(entry_num==3):
				t3_entry.delete(0, END)
				t3_entry.insert(END, t3)

	def numpad():
		global numpad_labelframe
		numpad_labelframe = LabelFrame(keypad_labelframe, bg="white", width=385, height=395)
		numpad_labelframe.place(x=2,y=1)
		button_list = ['7',     '8',      '9',
					   '4',     '5',      '6',
					   '1',     '2',      '3',
					   '0',     'Delete', 'Default']
		r = 1
		c = 0
		n = 0
		btn = list(range(len(button_list)))
		for label in button_list:
			cmd = partial(numpad_click, label)
			btn[n] = Button(numpad_labelframe, text=label, font=font.Font(family='Helvetica', size=10, weight='bold'), width=9, height=3, command=cmd)
			#btn[n] = atk.Button3d(numpad_labelframe, bg ='Lavender', text=label, width=15, command=cmd)
			btn[n].grid(row=r, column=c, padx=1, pady=1)
			n += 1
			c += 1
			if (c == 3):
				c = 0
				r += 1

	cir_img = Image.open('cir.png')
	cir_width, cir_height = cir_img.size
	scale_percent = 13
	width = int(cir_width * scale_percent / 100)
	height = int(cir_height * scale_percent / 100)
	display_img = cir_img.resize((width,height))
	image_select = ImageTk.PhotoImage(display_img)
	t1cir_label = Label(settemptop_labelframe, bg='white', image=image_select)
	t1cir_label.image = image_select
	t1cir_label.place(x=85,y=2)
	t2cir_label = Label(settemptop_labelframe, bg='white', image=image_select)
	t2cir_label.image = image_select
	t2cir_label.place(x=290,y=2)
	t3cir_label = Label(settemptop_labelframe, bg='white', image=image_select)
	t3cir_label.image = image_select
	t3cir_label.place(x=85,y=145)
	graycir_img = Image.open('graycir.png')
	graycir_width, cir_height = cir_img.size
	scale_percent = 13
	width = int(cir_width * scale_percent / 100)
	height = int(cir_height * scale_percent / 100)
	display_img = graycir_img.resize((width,height))
	image_select = ImageTk.PhotoImage(display_img)
	t4cir_label = Label(settemptop_labelframe, bg='white', image=image_select)
	t4cir_label.image = image_select
	t4cir_label.place(x=290,y=145)

	def entryt1_click(event):
		global numpad_labelframe
		global entry_num
		entry_num = 1
		numpad()
	def entryt2_click(event):
		global numpad_labelframe
		global entry_num
		entry_num = 2
		numpad()
	def entryt3_click(event):
		global numpad_labelframe
		global entry_num
		entry_num = 3
		numpad()

	t1_label = Label(settemptop_labelframe, bg='white', text='T1', fg='black', font=("Courier",17,"bold"))
	t1_label.place(x=90, y=15)
	t1oc_label = Label(settemptop_labelframe, bg='white', text=chr(176)+'C', fg='red', font=("Courier", 12,"bold"))
	t1oc_label.place(x=177, y=50)
	t1_entry = Entry(settemptop_labelframe, width=2, justify='center', bg='white', borderwidth=0, fg ='grey32', font=("Courier",31,"bold"))
	t1_entry.place(x=127,y=49)
	t1_entry.bind('<Button-1>', entryt1_click)
	global t1
	t1_entry.insert(0,t1)
	t2_label = Label(settemptop_labelframe, bg='white', text='T2', fg='black', font=("Courier",17,"bold"))
	t2_label.place(x=294, y=15)
	t2oc_label = Label(settemptop_labelframe, bg='white', text=chr(176)+'C', fg='red', font=("Courier", 12,"bold"))
	t2oc_label.place(x=381, y=50)
	t2_entry = Entry(settemptop_labelframe, width=2, justify='center', bg='white', borderwidth=0, fg ='grey32', font=("Courier",31,"bold"))
	t2_entry.place(x=331,y=49)
	t2_entry.bind('<Button-1>', entryt2_click)
	global t2
	t2_entry.insert(0,t2)
	t3_label = Label(settemptop_labelframe, bg='white', text='T3', fg='black', font=("Courier",17,"bold"))
	t3_label.place(x=90, y=158)
	t3oc_label = Label(settemptop_labelframe, bg='white', text=chr(176)+'C', fg='red', font=("Courier", 12,"bold"))
	t3oc_label.place(x=177, y=193)
	t3_entry = Entry(settemptop_labelframe, width=2, justify='center', bg='white', borderwidth=0, fg ='grey32', font=("Courier",31,"bold"))
	t3_entry.place(x=127,y=192)
	t3_entry.bind('<Button-1>', entryt3_click)
	global t3
	t3_entry.insert(0,t3)
	t4_label = Label(settemptop_labelframe, bg = 'white', text='T4', fg='grey67', font=("Courier",17,"bold"))
	t4_label.place(x=294, y=158)

	def back_click():
		settemp_labelframe.place_forget()
		mainscreen()
	def thread():
		th1 = Thread(target = next_click)
		th1.start()
	def next_click():
		settemp_labelframe.place_forget()
		global t1_set, t2_set, t3_set
		t1_set = t1_entry.get()
		t2_set = t2_entry.get()
		t3_set = t3_entry.get()
		scanposition()

	back_button = Button(settemp_labelframe, bg="Lavender", text="Back" , height=3, width=15, borderwidth=0, command=back_click)
	back_button.place(x=10,y=351)
	next_button = Button(settemp_labelframe, bg="Lavender", text="Next", height=3, width=15, borderwidth=0, command=thread)
	next_button.place(x=724,y=351)
	save_button = Button(settemp_labelframe, bg="yellow", text="Save", height=3, width=15, borderwidth=0)
	save_button.place(x=370,y=351)

##################################################### Giao diện định vị mẫu ############################################################   
def scanposition():
    global scanpostion_labelframe
    scanposition_labelframe = LabelFrame(root, bg='white', width=850, height=417)
    scanposition_labelframe.place(x=0,y=0)
    scanposition_label = Label(scanposition_labelframe, bg='white', text='SAMPLES POSITION', font=("Courier",17,'bold'), width=20, height=1 )
    scanposition_label.place(x=280,y=7)

    scan_img = Image.open('scan.png')
    scan_width, scan_height = scan_img.size
    scale_percent = 100
    width = int(scan_width * scale_percent / 100)
    height = int(scan_height * scale_percent / 100)
    display_img = scan_img.resize((width,height))
    image_select = ImageTk.PhotoImage(display_img)
    scan_label = Label(scanposition_labelframe, bg='white',image=image_select)
    scan_label.image = image_select
    scan_label.place(x=290,y=60)
    
    s = ttk.Style()
    s.theme_use('clam')
    s.configure("green.Horizontal.TProgressbar", foreground='green', background='green')
    scanposition_progressbar = ttk.Progressbar(root, orient = HORIZONTAL, style="green.Horizontal.TProgressbar", length = 200, mode = 'determinate')
    scanposition_progressbar.place(x=320,y=350)

    def thread():
        th1 = Thread(target = next_click)
        th1.start()
    def back_click():
        scanposition_labelframe.place_forget()
        settemp()   
    def next_click():
        scanposition_labelframe.place_forget()   
        analysis()
    back_button = Button(scanposition_labelframe, bg="Lavender", text="Back" , height=3, width=15, borderwidth=0, command=back_click)
    back_button.place(x=10,y=351)
    next_button = Button(scanposition_labelframe, bg="Lavender", text="Next", height=3, width=15, borderwidth=0,command=thread)
    next_button.place(x=724,y=351)

    global ser
    send_data = 'P'
    ser.write(send_data.encode())

    if(ser.in_waiting>0):
        receive_data = ser.readline().decode('utf-8').rstrip()
        print("Data received:", receive_data)
        scanposition_progressbar['value'] = 5
        root.update_idletasks()
        if(receive_data=='C'):
            global wait
            wait = 1
            scanposition_progressbar['value'] = 20
            root.update_idletasks()
            directory = strftime("COVID19 %y-%m-%d %H.%M.%S")
            global path0
            path0 = os.path.join("/home/pi/Desktop/spotcheck result",directory)
            os.mkdir(path0)
            global path1
            path1 = os.path.join(path0,"Original image")
            os.mkdir(path1)
            global path2
            path2 = os.path.join(path0,"Processed image")
            os.mkdir(path2)
            global path3
            path3 = os.path.join(path0,"Result Table")
            os.mkdir(path3)
            global path4
            path4 = os.path.join(path0,"Sample image")
            os.mkdir(path4)
            global path5
            path5 = os.path.join(path0,"Temperature program")
            os.mkdir(path5)
    while(wait!=1):
        scanposition_progressbar['value'] = 2
        root.update_idletasks()
        if(ser.in_waiting>0):
            receive_data = ser.readline().decode('utf-8').rstrip()
            print("Data received:", receive_data)
            scanposition_progressbar['value'] = 10
            root.update_idletasks()
            if(receive_data=='C'):
                scanposition_progressbar['value'] = 20
                root.update_idletasks()
                directory = strftime("COVID19 %y-%m-%d %H.%M.%S")
                path0 = os.path.join("/home/pi/Desktop/spotcheck result",directory)
                os.mkdir(path0)
                path1 = os.path.join(path0,"Original image")
                os.mkdir(path1)
                path2 = os.path.join(path0,"Processed image")
                os.mkdir(path2)
                path3 = os.path.join(path0,"Result Table")
                os.mkdir(path3)
                path4 = os.path.join(path0,"Sample image")
                os.mkdir(path4)
                path5 = os.path.join(path0,"Temperature program")
                os.mkdir(path5)
                wait = 1
                break;
    while(wait==1):
        camera = PiCamera(framerate=Fraction(1,6), sensor_mode=3)
        camera.rotation = 180
        camera.iso = 800
        sleep(2)
        camera.shutter_speed = 6000000
        camera.exposure_mode = 'off'
        output = path1 + "/Sample.jpg"
        camera.capture(output)
        camera.close()
        scanposition_progressbar['value'] = 40
        root.update_idletasks()

        global pos_result
        pos_result, pos_image,__,__ = process_image(output)
        scanposition_progressbar['value'] = 60
        root.update_idletasks()
        sleep(1)

        output = path4 + "/Sample.jpg"
        cv2.imwrite(output, pos_image)
        scanposition_progressbar['value'] = 90
        root.update_idletasks()
        sleep(1)

        scanresult_labelframe = LabelFrame(scanposition_labelframe, bg='ghost white', width=528,height = 307)
        scanresult_labelframe.place(x=245,y=42)

        label = list(range(48))
        def result_table(range_a, range_b, row_value):
            j=-1
            for i in range(range_a, range_b):
                j+=1
                if(i<6):
                    t='A'+ str(i+1)
                if(i>=6 and i<12):
                    t='B'+ str(i-5)
                if(i>=12 and i<18):
                    t='C'+ str(i-11)
                if(i>=18 and i<24):
                    t='D'+ str(i-17)
                if(i>=24 and i<30):
                    t='E'+ str(i-23)
                if(i>=30 and i<36):
                    t='F'+ str(i-29)
                if(i>=36 and i<42):
                    t='G'+ str(i-35)
                if(i>=42):
                    t='H'+ str(i-41)
                if(pos_result[i]< 12):
                    label[i] = Label(scanresult_labelframe, bg='gainsboro', text=t, width=5, height=2)
                    label[i].grid(row=row_value,column=j,padx=3,pady=3)
                else:
                    label[i] = Label(scanresult_labelframe, bg='dodger blue', text=t, width=5, height=2)
                    label[i].grid(row=row_value,column=j,padx=3,pady=3) 
        scanposition_progressbar['value'] = 100
        root.update_idletasks() 

        result_table(0,6,0)
        result_table(6,12,1)
        result_table(12,18,2)
        result_table(18,24,3)
        result_table(24,30,4)
        result_table(30,36,5)
        result_table(36,42,6)
        result_table(42,48,7)
        scan_label.place_forget()
        scanposition_progressbar.place_forget()
        wait = 0

################################################### Giao diện phân tích mẫu ###########################################################
def analysis():
    global analysis_labelframe
    analysis_labelframe = LabelFrame(root, bg='white', width=850, height=417)
    analysis_labelframe.place(x=0,y=0)
    analysis_label = Label(analysis_labelframe, bg='white', text='SAMPLES ANALYSIS', font=("Courier",17,'bold'), width=20, height=1 )
    analysis_label.place(x=280,y=7)
    t_labelframe = LabelFrame(analysis_labelframe, bg='white', width=846, height=298)
    t_labelframe.place(x=0,y=40)

    t1_labelframe = LabelFrame(analysis_labelframe, bg='white',text="T1:"+t1_set+chr(176)+'C' , font=("Courier",13,'bold'), width=200, height=280)
    t1_labelframe.place(x=5,y=45)
    t2_labelframe = LabelFrame(analysis_labelframe, bg='white',text="T2:"+t2_set+chr(176)+'C' , font=("Courier",13,'bold'), width=200, height=280)
    t2_labelframe.place(x=217,y=45)
    t3_labelframe = LabelFrame(analysis_labelframe, bg='white',text="T1:"+t3_set+chr(176)+'C' , font=("Courier",13,'bold'), width=200, height=280)
    t3_labelframe.place(x=429,y=45)
    t4_labelframe = LabelFrame(analysis_labelframe, bg='white smoke',text="T4", width=200, height=280)
    t4_labelframe.place(x=641,y=45)
    autoprocess_label = Label(analysis_labelframe, bg='white', text="Program is automatically processing...", fg='blue', font=("Courier",8,'bold'))
    autoprocess_label.place(x=550,y=365)
    temp_label = Label(analysis_labelframe, bg='white', fg='grey36', text='Temperature: fsfsC', font=("Courier",17,'bold'))
    temp_label.place(x=32,y=359)

    def stop_click():
        global ser
        msgbox = messagebox.askquestion('Stop the process','Are you want to stop the analysis ?', icon = 'warning')
        if(msgbox=='yes'):
            send_data ='S'
            ser.write(send_data.encode())
            analysis_labelframe.place_forget()
            mainscreen()       
    
    def pause_click():
        global ser
        if(pause_button['text']=='Pause'):
            send_data ='P'
            ser.write(send_data.encode())
            pause_button['text']= 'Continue'
        else:
            send_data ='R'
            ser.write(send_data.encode())
            pause_button['text']= 'Pause'
    pause_button = Button(analysis_labelframe, bg="Lavender", text="Pause" , height=4, width=12, borderwidth=0)
    pause_button.place(x=323,y=342)
    stop_button = Button(analysis_labelframe, bg="Lavender", text="Stop", height=4, width=12, borderwidth=0)
    stop_button.place(x=432,y=342)
    root.update()

    global t1_set, t2_set, t3_set
    send_data = "t"+ t1_set + "," + t2_set + "," + t3_set + "z"
    global ser
    ser.write(send_data.encode())
    t0 = time.time()
    sleep(2)
    #global ser
    global wait
    if(ser.in_waiting>0):
        receive_data = ser.readline().decode('ascii').rstrip()
        print("Data received:", receive_data)        
        if(receive_data=='Y'):
            wait = 1

    while(wait!=1):
        if(ser.in_waiting>0):
            receive_data = ser.readline().decode('ascii').rstrip()
            print("Data received:", receive_data)        
            if(receive_data=='Y'):
                wait = 1
               	break

	while(wait==1):
		if(ser.in_waiting>0):
            receive_data = ser.readline().decode('ascii').rstrip()
            #print("Data received:", receive_data)
            if(receive_data!='C1' and receive_data!='C2' and receive_data!='C3'):
                print("Data received:", receive_data)
                temp_label['text'] = 'Temperature: '+ receive_data+ ' ' + chr(176)+'C'
                root.update()
                
            if(receive_data=='C1'):
                print("Data received:", receive_data)
                t1_labelframe['bg'] = atk.DEFAULT_COLOR
                t1_labelframe['fg'] = 'lawn green'
                t_progressbar = atk.RadialProgressbar(t1_labelframe, fg='cyan')
                t_progressbar.place(x=45,y=73)
                t_progressbar.start()
                tprocess_label = Label(t1_labelframe, bg=atk.DEFAULT_COLOR, fg='white smoke', text='Processing!', font=("Courier",9,'bold'))
                tprocess_label.place(x=54,y=112)

                camera = PiCamera(framerate=Fraction(1,6), sensor_mode=3)
                camera.rotation = 180
                camera.iso = 800
                sleep(2)
                camera.shutter_speed = 6000000
                camera.exposure_mode = 'off'

                global path1
                output1 = path1 + "/T1.jpg"
                camera.capture(output1)
                camera.close()
                send_data = 'C'
                ser.write(send_data.encode())
                print('Capture done!')

                t1_result, t1_image, t1_start, t1_end = process_image(output1)

                global path2
                output = path2 + "/T1.jpg"
                cv2.imwrite(output, t1_image)

                t1_analysis = Image.open(output)
                t1_crop = t1_analysis.crop((t1_start[0]-7,t1_start[1]-7,t1_end[0]+7,t1_end[1]+7))   
                crop_width, crop_height = t1_crop.size
                scale_percent = 95
                width = int(crop_width * scale_percent / 100)
                height = int(crop_height * scale_percent / 100)
                display_img = t1_crop.resize((width,height))
                t1_display = ImageTk.PhotoImage(display_img)                
                t1_label = Label(t1_labelframe, image=t1_display)
                t1_label.image = t1_display
                t1_label.place(x=0,y=0)
                root.update()

                sheet["A2"] = "A"
                sheet["A3"] = "B"
                sheet["A4"] = "C"
                sheet["A5"] = "D"
                sheet["A6"] = "E"
                sheet["A7"] = "F"
                sheet["A8"] = "G"
                sheet["A9"] = "H"
                sheet["B1"] = "1"
                sheet["C1"] = "2"
                sheet["D1"] = "3"
                sheet["E1"] = "4"
                sheet["F1"] = "5"
                sheet["G1"] = "6"
                for i in range(0,48):
                    if(i<6):
                        pos = str(chr(65+i+1)) + "2"
                    if(i>=6 and i<12):
                        pos = str(chr(65+i-5)) + "3"
                    if(i>=12 and i<18):
                        pos = str(chr(65+i-11)) + "4"
                    if(i>=18 and i<24):
                        pos = str(chr(65+i-17)) + "5"
                    if(i>=24 and i<30):
                        pos = str(chr(65+i-23)) + "6"
                    if(i>=30 and i<36):
                        pos = str(chr(65+i-29)) + "7"
                    if(i>=36 and i<42):
                        pos = str(chr(65+i-35)) + "8"
                    if(i>=42):
                        pos = str(chr(65+i-41)) + "9"
                    
                    sheet[pos] = t1_result[i]
                
                global path3
                workbook.save(path3+"/T1.xlsx")

            if(receive_data=='C2'):
                print("Data received:", receive_data)
                t2_labelframe['bg'] = atk.DEFAULT_COLOR
                t2_labelframe['fg'] = 'lawn green'
                t_progressbar = atk.RadialProgressbar(t2_labelframe, fg='cyan')
                t_progressbar.place(x=45,y=73)
                t_progressbar.start()
                tprocess_label = Label(t2_labelframe, bg=atk.DEFAULT_COLOR, fg='white smoke', text='Processing!', font=("Courier",9,'bold'))
                tprocess_label.place(x=54,y=112)

                camera = PiCamera(framerate=Fraction(1,6), sensor_mode=3)
                camera.rotation = 180
                camera.iso = 800
                sleep(2)
                camera.shutter_speed = 6000000
                camera.exposure_mode = 'off'
                output2 = path1 + "/T2.jpg"
                camera.capture(output2)
                camera.close()
                send_data = 'C'
                ser.write(send_data.encode())
                print('Capture done!')
                t2_result, t2_image, t2_start, t2_end = process_image(output2)
                output = path2 + "/T2.jpg"
                cv2.imwrite(output, t2_image)
                t2_analysis = Image.open(output)
                t2_crop = t2_analysis.crop((t2_start[0]-7,t2_start[1]-7,t2_end[0]+7,t2_end[1]+7))

                crop_width, crop_height = t2_crop.size
                scale_percent = 95
                width = int(crop_width * scale_percent / 100)
                height = int(crop_height * scale_percent / 100)
                display_img = t2_crop.resize((width,height))
                t2_display = ImageTk.PhotoImage(display_img)
                
                #t2_display = ImageTk.PhotoImage(t2_crop)
                t2_label = Label(t2_labelframe, image=t2_display)
                t2_label.image = t2_display
                t2_label.place(x=0,y=0)
                root.update()

                sheet["A2"] = "A"
                sheet["A3"] = "B"
                sheet["A4"] = "C"
                sheet["A5"] = "D"
                sheet["A6"] = "E"
                sheet["A7"] = "F"
                sheet["A8"] = "G"
                sheet["A9"] = "H"
                sheet["B1"] = "1"
                sheet["C1"] = "2"
                sheet["D1"] = "3"
                sheet["E1"] = "4"
                sheet["F1"] = "5"
                sheet["G1"] = "6"
                for i in range(0,48):
                    if(i<6):
                        pos = str(chr(65+i+1)) + "2"
                    if(i>=6 and i<12):
                        pos = str(chr(65+i-5)) + "3"
                    if(i>=12 and i<18):
                        pos = str(chr(65+i-11)) + "4"
                    if(i>=18 and i<24):
                        pos = str(chr(65+i-17)) + "5"
                    if(i>=24 and i<30):
                        pos = str(chr(65+i-23)) + "6"
                    if(i>=30 and i<36):
                        pos = str(chr(65+i-29)) + "7"
                    if(i>=36 and i<42):
                        pos = str(chr(65+i-35)) + "8"
                    if(i>=42):
                        pos = str(chr(65+i-41)) + "9"
                    
                    sheet[pos] = t2_result[i]
                
                workbook.save(path3+"/T2.xlsx")
    
            if(receive_data=='C3'):
                print("Data received:", receive_data)
                t3_labelframe['bg'] = atk.DEFAULT_COLOR
                t3_labelframe['fg'] = 'lawn green'
                t_progressbar = atk.RadialProgressbar(t3_labelframe, fg='cyan')
                t_progressbar.place(x=45,y=73)
                t_progressbar.start()
                tprocess_label = Label(t3_labelframe, bg=atk.DEFAULT_COLOR, fg='white smoke', text='Processing!', font=("Courier",9,'bold'))
                tprocess_label.place(x=54,y=112)

                camera = PiCamera(framerate=Fraction(1,6), sensor_mode=3)
                camera.rotation = 180
                camera.iso = 800
                sleep(2)
                camera.shutter_speed = 6000000
                camera.exposure_mode = 'off'
                output3 = path1 + "/T3.jpg"
                camera.capture(output3)
                camera.close()
                send_data = 'C'
                ser.write(send_data.encode())
                print('Capture done!')
                t3_result, t3_image, t3_start, t3_end = process_image(output3)
                output = path2 + "/T3.jpg"
                cv2.imwrite(output, t3_image)
                t3_analysis = Image.open(output)
                t3_crop = t3_analysis.crop((t3_start[0]-7,t3_start[1]-7,t3_end[0]+7,t3_end[1]+7))
                
                crop_width, crop_height = t3_crop.size
                scale_percent = 95
                width = int(crop_width * scale_percent / 100)
                height = int(crop_height * scale_percent / 100)
                display_img = t3_crop.resize((width,height))
                t3_display = ImageTk.PhotoImage(display_img)
                
                t3_label = Label(t3_labelframe, image=t3_display)
                t3_label.image = t3_display
                t3_label.place(x=0,y=0)
                wait = 0
                root.update()
                pause_button.place_forget()
                stop_button.place_forget()
                temp_label.place_forget()
                autoprocess_label.place_forget()

                sheet["A2"] = "A"
                sheet["A3"] = "B"
                sheet["A4"] = "C"
                sheet["A5"] = "D"
                sheet["A6"] = "E"
                sheet["A7"] = "F"
                sheet["A8"] = "G"
                sheet["A9"] = "H"
                sheet["B1"] = "1"
                sheet["C1"] = "2"
                sheet["D1"] = "3"
                sheet["E1"] = "4"
                sheet["F1"] = "5"
                sheet["G1"] = "6"
                for i in range(0,48):
                    if(i<6):
                        pos = str(chr(65+i+1)) + "2"
                    if(i>=6 and i<12):
                        pos = str(chr(65+i-5)) + "3"
                    if(i>=12 and i<18):
                        pos = str(chr(65+i-11)) + "4"
                    if(i>=18 and i<24):
                        pos = str(chr(65+i-17)) + "5"
                    if(i>=24 and i<30):
                        pos = str(chr(65+i-23)) + "6"
                    if(i>=30 and i<36):
                        pos = str(chr(65+i-29)) + "7"
                    if(i>=36 and i<42):
                        pos = str(chr(65+i-35)) + "8"
                    if(i>=42):
                        pos = str(chr(65+i-41)) + "9"
                    
                    sheet[pos] = t3_result[i]
                
                workbook.save(path3+"/T3.xlsx")
                
                def viewresult_click():
                    viewresult_button.place_forget()
                    t1_labelframe.place_forget()
                    t2_labelframe.place_forget()
                    t3_labelframe.place_forget()

                    annotate_labelframe = LabelFrame(analysis_labelframe, bg='white', width=395, height=385)
                    annotate_labelframe.place(x=360,y=11)

                    def finish_click():
                        msgbox = messagebox.askquestion('Finish','Are you want to return to the main menu ?', icon = 'warning')
                        if(msgbox=='yes'):
                            analysis_labelframe.place_forget()
                            mainscreen() 
                    finish_button = Button(annotate_labelframe, bg="lavender", text="FINISH", height=3, width=15, borderwidth=0, command=finish_click)
                    finish_button.place(x=137,y=323)

                    result_labelframe = LabelFrame(analysis_labelframe, bg='ghost white', width=600,height = 307)
                    result_labelframe.place(x=104,y=56)
                    row_labelframe = LabelFrame(analysis_labelframe, bg='ghost white', width=600,height = 50)
                    row_labelframe.place(x=104,y=12)
                    column_labelframe = LabelFrame(analysis_labelframe, bg='ghost white', width=50,height = 307)
                    column_labelframe.place(x=62,y=56)
                                                    
                    row_label = [0,0,0,0,0,0]
                    for i in range (0,6):
                        row_text = i+1
                        row_label[i] = Label(row_labelframe, text=row_text, bg='white smoke', width=4, height=2)
                        row_label[i].grid(row=0,column=i,padx=2,pady=2)
                                                        
                    column_label = [0,0,0,0,0,0,0,0]
                    for i in range (0,8):
                        if(i==0):
                            column_text = 'A'
                        if(i==1):
                            column_text = 'B'
                        if(i==2):
                            column_text = 'C'
                        if(i==3):
                            column_text = 'D'
                        if(i==4):
                            column_text = 'E'
                        if(i==5):
                            column_text = 'F'
                        if(i==6):
                            column_text = 'G'
                        if(i==7):
                            column_text = 'H'
                        column_label[i] = Label(column_labelframe, text=column_text, bg='white smoke', width=4, height=2)
                        column_label[i].grid(row=i,column=0,padx=2,pady=2)

                    label = list(range(48))
                    def result_table(range_a,range_b, row_value):
                        j=-1
                        global pos_result
                        for i in range(range_a, range_b):
                            j+=1
                            if(pos_result[i]<12):
                                label[i] = Label(result_labelframe, bg='white smoke', text=str('%d'%pos_result[i]), width=4, height=2)
                                label[i].grid(row=row_value,column=j,padx=2,pady=2)
                            else:
                                if(t1_result[i]<20):
                                    label[i] = Label(result_labelframe, bg='yellow', text=str('%d'%t3_result[i]), width=4, height=2)
                                    label[i].grid(row=row_value,column=j,padx=2,pady=2)
                                if(t1_result[i]>=20 and t2_result[i]<20):
                                    label[i] = Label(result_labelframe, bg='yellow', text=str('%d'%t3_result[i]), width=4, height=2)
                                    label[i].grid(row=row_value,column=j,padx=2,pady=2) 
                                if(t1_result[i]>=20 and t2_result[i]>=20 and t3_result[i]<20):
                                    label[i] = Label(result_labelframe, bg='lawn green', text=str('%d'%t3_result[i]), width=4, height=2)
                                    label[i].grid(row=row_value,column=j,padx=2,pady=2)
                                if(t1_result[i]>=20 and t2_result[i]>=20 and t3_result[i]>20 and t3_result[i]<22):
                                    label[i] = Label(result_labelframe, bg='cyan', text=str('%d'%t3_result[i]), width=4, height=2)
                                    label[i].grid(row=row_value,column=j,padx=2,pady=2)
                                if(t1_result[i]>=20 and t2_result[i]>=20 and t3_result[i]>=22):
                                    label[i] = Label(result_labelframe, bg='red', text=str('%d'%t3_result[i]), width=4, height=2)
                                    label[i].grid(row=row_value,column=j,padx=2,pady=2)
                                                    
                    result_table(0,6,0)
                    result_table(6,12,1)
                    result_table(12,18,2)
                    result_table(18,24,3)
                    result_table(24,30,4)
                    result_table(30,36,5)
                    result_table(36,42,6)
                    result_table(42,48,7)

                viewresult_button = Button(analysis_labelframe, bg="lavender", text="VIEW RESULT", height=3, width=15, borderwidth=0, command=viewresult_click)
                viewresult_button.place(x=370,y=351)


########################################################## Biến toàn cục ################################################################# 
covid19clicked = 0
tbclicked = 0
viewresultclicked = 0
t1 = '30'
t2 = '50'
t3 = '73'
t1_set = t1
t2_set = t2
t3_set = t3
t1_set = t1
t2_set = t2
t3_set = t3
temp_label = 0
name = "/"
entry_num = 0
wait = 0
pos_result = list(range(48))
path0 = "/"
path1 = "/"
path2 = "/"
path3 = "/"
path4 = "/"
path5 = "/"

########################################################## Khoi tao excel ###############################################################
workbook = Workbook()
sheet = workbook.active

########################################################### Serial init ################################################################ 
# ser = serial.Serial('/dev/ttyACM0', 9600)
ser = serial.Serial(
    port = '/dev/serial0',
    baudrate = 115200,
    parity = serial.PARITY_NONE,
    stopbits = serial.STOPBITS_ONE,
    bytesize = serial.EIGHTBITS,
    timeout = 1
)

############################################################## Loop ####################################################################
while True:
	mainscreen()  
	root.mainloop()