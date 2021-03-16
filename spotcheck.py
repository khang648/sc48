########################################################## Import Module ################################################################
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import time
from time import sleep, gmtime, strftime
from picamera import PiCamera
import cv2 
import numpy as np
from tkinter import filedialog as fd
from PIL import ImageTk, Image
import serial
from functools import partial
import math
from fractions import Fraction
from threading import *
import os
from openpyxl import Workbook
#from gpiozero import LED

######################################################### Khởi tạo Main Window ##############################################################
root = Tk()
root.title("SpotCheck Application")
root.geometry('850x450')
root.configure(background = "white")

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
# new led   
#     for i in range(len(sorted_contours1)):
#         if(i==0 or i==1):
#             result_list[i] *= 1.4
#         if(i==2 or i==4 or i==11 or i==41 or i==46 or i==47):
#             result_list[i] *= 1.3
#         if(i==3 or i==6 or i==7 or i==12 or i==13 or i==17 or i==18 or i==19 or i==23
#            or i==24 or i==25 or i==29 or i==30 or i==31 or i==35 or i==36 or i==37
#            or i==40 or i==42 or i==43 or i==44 or i==45):
#             result_list[i] *= 1.2
#         if(i==8 or i==9 or i==10 or i==14 or i==15 or i==16 or i==22 or i==28
#            or i==32 or i==33 or i==34 or i==38 or i==39):
#             result_list[i] *= 1.1
#         if(i==5):
#             result_list[i] *= 1.5
    
# new led:
    for i in range(len(sorted_contours1)):
        if(i==0 or i==1 or i==4 or i==11):
            result_list[i] *= 1.24
        if(i==2 or i==3 or i==7 or i==6 or i==12 or i==38 or i==46 or i==47):
            result_list[i] *= 1.17
        if(i==5):
            result_list[i] *= 1.4
        if(i==9 or i==10 or i==13 or i==16 or i==17 or i==18 or i==23
           or i==30 or i==31 or i==36 or i==37 or i==39 or i==40 or i==41
           or i==42 or i==43 or i==44 or i==45):
            result_list[i] *= 1.11
        if(i==14 or i==15 or i==19 or i==22 or i==24 or i==29 or i==32
           or i==33 or i==34):
            result_list[i] *= 1.05
            
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
    global mainscreen_labelframe
    mainscreen_labelframe = LabelFrame(root, bg='white', width=850, height=417)
    mainscreen_labelframe.place(x=0,y=0)

    def newprogram_click():
        mainscreen_labelframe.place_forget()
        settemp()

    newprogram_button = Button(mainscreen_labelframe, bg="lavender", text="NEW PROGRAM", height=4, width=20, command=newprogram_click)
    newprogram_button.place(x=300,y=100)
    covid19_button = Button(mainscreen_labelframe, bg="lavender", text="COVID 19", height=4, width=20)
    covid19_button.place(x=300,y=200)
    viewresult_button = Button(mainscreen_labelframe, bg="lavender", text="VIEW RESULT", height=4, width=20)
    viewresult_button.place(x=300,y=300)

 ###################################################### Giao diện set nhiệt độ ############################################################ 
def settemp():
    global settemp1_labelframe
    global settemp11_labelframe
    settemp1_labelframe = LabelFrame(root, bg='white', width=850, height=417)
    settemp1_labelframe.place(x=0,y=0) 
    settemp_label = Label(settemp1_labelframe, bg='white', text='SET TEMPERATURE', font=("Courier",20), width=20, height=1 )
    settemp_label.place(x=235,y=0)
    settemp11_labelframe = LabelFrame(settemp1_labelframe,bg='white',width=675, height=301)
    settemp11_labelframe.place(x=60,y=40)
    settemp111_labelframe = LabelFrame(settemp11_labelframe,bg='white',width=400, height=300)
    settemp111_labelframe.place(x=1,y=1)    
    t1_label = Label(settemp111_labelframe, bg='orange', anchor='e', text='T1:', font=("Courier",20), width=14, height=3)
    t1_label.grid(row=0,column=1)
    t2_label = Label(settemp111_labelframe, bg='dark orange', anchor='e', text='T2:', font=("Courier",20), width=14, height=3)
    t2_label.grid(row=1,column=1)
    t3_label = Label(settemp111_labelframe, bg='tomato', anchor='e', text='T3:', font=("Courier",20), width=14, height=3)
    t3_label.grid(row=2,column=1)

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
            #root.unbind('<Button-1>', btn_funcid)

    def numpad():
        global numpad_labelframe
        numpad_labelframe = LabelFrame(settemp11_labelframe, bg="white", width=500, height=300)
        numpad_labelframe.place(x=390,y=1)
        button_list = [
            '7',     '8',      '9',
            '4',     '5',      '6',
            '1',     '2',      '3',
            '0',     'Delete', 'Default']
        r = 1
        c = 0
        n = 0
        btn = list(range(len(button_list)))
        for label in button_list:
            cmd = partial(numpad_click, label)
            btn[n] = Button(numpad_labelframe, text=label, width=7, height=3, command=cmd)
            btn[n].grid(row=r, column=c, padx=4, pady=4)
            n += 1
            c += 1
            if (c == 3):
                c = 0
                r += 1

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

    t1_entry = Entry(settemp111_labelframe, width=4, justify='center', font=("Courier",25))
    t1_entry.grid(row=0,column=2,padx=8,pady=8)
    t1_entry.bind('<Button-1>', entryt1_click)
    global t1
    t1_entry.insert(0,t1)

    t2_entry = Entry(settemp111_labelframe, width=4, justify='center', font=("Courier",25))
    t2_entry.grid(row=1,column=2,padx=8,pady=8)
    t2_entry.bind('<Button-1>', entryt2_click)
    global t2
    t2_entry.insert(0,t2)

    t3_entry = Entry(settemp111_labelframe, width=4, justify='center', font=("Courier",25))
    t3_entry.grid(row=2,column=2,padx=8,pady=8)
    t3_entry.bind('<Button-1>', entryt3_click)
    global t3
    t3_entry.insert(0,t3)
     
    t1_label = Label(settemp111_labelframe, bg='white', anchor='w', text=chr(176)+'C', font=("Courier",20), width=3)
    t1_label.grid(row=0,column=3,padx=6,pady=8)
    t2_label = Label(settemp111_labelframe, bg='white', anchor='w', text=chr(176)+'C', font=("Courier",20), width=3)
    t2_label.grid(row=1,column=3,padx=6,pady=8)
    t3_label = Label(settemp111_labelframe, bg='white', anchor='w', text=chr(176)+'C', font=("Courier",20), width=3)
    t3_label.grid(row=2,column=3,padx=6,pady=8)
    
    def thread():
        th1 = Thread(target = next_click)
        th1.start()
    def back_click():
        settemp1_labelframe.place_forget()
        mainscreen()
    
    def next_click():
        settemp1_labelframe.place_forget()
        global t1_set, t2_set, t3_set
        t1_set = t1_entry.get()
        t2_set = t2_entry.get()
        t3_set = t3_entry.get()
        scanposition()
                 
    back_button = Button(settemp1_labelframe, bg="lavender", text="<< Back", height=2, width=10, command=back_click)
    back_button.place(x=40,y=350)
    next_button = Button(settemp1_labelframe, bg="lavender", text="Next >>", height=2, width=10, command=thread)
    next_button.place(x=650,y=350)
    save_button = Button(settemp1_labelframe, bg="lavender", text="Save", height=2, width=10)
    save_button.place(x=345,y=350)
 
##################################################### Giao diện định vị mẫu ############################################################   
def scanposition():
    
    global scanpostion_labelframe
    scanposition_labelframe = LabelFrame(root, bg='white', width=850, height=417)
    scanposition_labelframe.place(x=0,y=0)
    scanposition_label = Label(scanposition_labelframe, bg='white', text='SAMPLES POSITION', font=("Courier",20), width=20, height=1 )
    scanposition_label.place(x=235,y=0)

    s = ttk.Style()
    s.theme_use('clam')
    s.configure("green.Horizontal.TProgressbar", foreground='green', background='green')
    scanposition_progressbar = ttk.Progressbar(root, orient = HORIZONTAL, style="green.Horizontal.TProgressbar", length = 150, mode = 'determinate')
    scanposition_progressbar.place(x=330,y=350)

    def thread():
        th1 = Thread(target = next_click)
        th1.start()
    def back_click():
        scanposition_labelframe.place_forget()
        settemp()   
    def next_click():
        scanposition_labelframe.place_forget()   
        analysis()
            
    back_button = Button(scanposition_labelframe, bg="lavender", text="<< Back", height=2, width=10, command=back_click)
    back_button.place(x=40,y=350)
    next_button = Button(scanposition_labelframe, bg="lavender", text="Next >>", height=2, width=10, command=thread)
    next_button.place(x=650,y=350)

    global ser
    ser.flushInput()
    ser.flushOutput()
    send_data = 'P'
    ser.write(send_data.encode())
    #global ser

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
        scanposition_progressbar.place_forget()
        wait = 0
      
################################################### Giao diện phân tích mẫu ###########################################################
def analysis():
    global analysis_labelframe
    analysis_labelframe = LabelFrame(root, bg='white', width=850, height=417)
    analysis_labelframe.place(x=0,y=0)
    
    t1_labelframe = LabelFrame(analysis_labelframe, bg='white smoke',text="T1", width=250, height=340)
    t1_labelframe.place(x=8,y=0)
    t2_labelframe = LabelFrame(analysis_labelframe, bg='white smoke',text="T2", width=250, height=340)
    t2_labelframe.place(x=274,y=0)
    t3_labelframe = LabelFrame(analysis_labelframe, bg='white smoke',text="T3", width=250, height=340)
    t3_labelframe.place(x=540,y=0)
    temp_label = Label(analysis_labelframe, bg='white')
    temp_label.place(x=340,y=350)

    def cancel_click():
        global ser
        msgbox = messagebox.askquestion('Cancel the process','Are you want to cancel the analysis ?', icon = 'warning')
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
    pause_button = Button(scanposition_labelframe, bg="lavender", text="Pause", height=2, width=10, command = pause_click)
    pause_button.place(x=40,y=350)
    cancel_button = Button(scanposition_labelframe, bg="lavender", text="Cancel", height=2, width=10, command=cancel_click)
    cancel_button.place(x=650,y=350)
    root.update()
    
    global t1_set, t2_set, t3_set
    send_data = "t"+ t1_set + "," + t2_set + "," + t3_set + "z"
    global ser
    ser.flushInput()
    ser.flushOutput()
    ser.write(send_data.encode())
    t0 = time.time()
    sleep(2)
    #global ser
    global wait
    if(ser.in_waiting>0):
        receive_data = ser.readline().decode('utf-8').rstrip()
        print("Data received:", receive_data)        
        if(receive_data=='Y'):
            wait = 1
    while(wait!=1):
        if(ser.in_waiting>0):
            receive_data = ser.readline().decode('utf-8').rstrip()
            print("Data received:", receive_data)        
            if(receive_data=='Y'):
                wait = 1
                break
        
#         sleep(3)
#         send_data = "t"+ t1_set + "," + t2_set + "," + t3_set + "z"
#         ser.flushOutput()
#         ser.write(send_data.encode())
#         print("send data again:", send_data)
#         if(ser.in_waiting>0):
#             receive_data = ser.readline().decode('ascii').rstrip()
#             print("Data received:", receive_data)        
#             if(receive_data=='Y'):   
#                 wait = 1
#                 break
    while(wait==1):
        if(ser.in_waiting>0):
            receive_data = ser.readline().decode('utf-8').rstrip()
            #print("Data received:", receive_data)
            if(receive_data!='C1' and receive_data!='C2' and receive_data!='C3'):
                print("Data received:", receive_data)
                temp_label['text'] = 'Temperature: '+ receive_data+ ' ' + chr(176)+'C'
                root.update()
                
            if(receive_data=='C1'):
                print("Data received:", receive_data)
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
                cancel_button.place_forget()
                temp_label.place_forget()

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
                    finish_button = Button(annotate_labelframe, bg="lavender", text="FINISH", height=2, width=10, command=finish_click)
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

                viewresult_button = Button(analysis_labelframe, bg="lavender", text="VIEW RESULT", height=2, width=10, command=viewresult_click)
                viewresult_button.place(x=345,y=350)

########################################################## Biến toàn cục ############################################################### 
mainscreen_labelframe = 0
settemp1_labelframe = 0
settemp11_labelframe = 0
scanposition_labelframe = 0
analysis_labelframe = 0
numpad_labelframe = 0
temp_label = 0
t1 = 30
t2 = 65
t3 = 80
t1_set = t1
t2_set = t2
t3_set = t3
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

######################################################### Khoi tao excel ###############################################################
workbook = Workbook()
sheet = workbook.active

########################################################### Serial init ################################################################ 
# ser = serial.Serial('/dev/ttyACM0', 9600)

# from gpiozero import LED
# tx = LED(14)
# rx = LED(15)
# rx.off()
# tx.off()
# sleep(1)

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