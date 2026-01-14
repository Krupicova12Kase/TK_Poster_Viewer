################################################################################
# Report issues at https://github.com/Krupicova12Kase/TK_Poster_Viewer
# Created by Krupicova12Kase, AKA Máťa or luki
# MIT license 
# Copyright (c) 2026 Krupicova12Kase
################################################################################

#settings
close_powerpoint = True #Should the program close powerpoint when it's done with generating? Not closing it may cause some problems
install_modules = True #Should the program install required packages? Packages are on line below
packages = ["pillow","pywin32"]

#Imports
import win32com.client
import os
from PIL import Image
import time
import subprocess

#Thank you stackoverflow https://stackoverflow.com/questions/287871/how-do-i-print-colored-text-to-the-terminal
class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    
#Module instalation and update
def update(module):
    subprocess.check_call(f'pip install {module}', shell=True)

if install_modules:
    try:
        for i in packages:
            update(i)
    except:
        print(f"{bcolors.WARNING}Failed to install packages!{bcolors.ENDC}")
    
names = [] 
passed = False

 
#Gemini helped with this powerpoint stuff   
def export_slide(ppt_app,pptx_path, output_folder,file):      
    # Open the presentation
    abs_path = os.path.abspath(pptx_path)
    presentation = ppt_app.Presentations.Open(abs_path, WithWindow=False)
    time.sleep(1)
    # Export the first slide (Index starts at 1)
    slide = presentation.Slides(1)
    output_path = os.path.join(os.path.abspath(output_folder), file)
    
    # Export method (FileName, FilterName, Width, Height)
    slide.Export(output_path, "PNG")
    print(f"Exported to: {bcolors.OKBLUE}{output_path}{bcolors.ENDC}")

    #Clean up
    presentation.Close()
    
directory = os.path.dirname(os.path.abspath(__file__))

#Check if files are valid
x = 0
for file in os.listdir(directory):
    filename = os.fsdecode(file)
    if filename.endswith(".pptx"): 
        x += 1

#Print error message when files are invalid
if x == 4:
    print(f"{bcolors.OKGREEN}Found {4} .pptx files, attempting conversion{bcolors.ENDC}")
    passed = True
    print("")
elif x > 4 or x < 4:
    print(f"{bcolors.FAIL}Found {x} .pptx files, unable to convert, need exactly 4!{bcolors.ENDC}")
    passed = False
    input()
else:
    print(f"{bcolors.FAIL}Something strange happened during checking amount of .pptx files, make sure there are exactly four! (found {x})If there are exactly four, open an issue on GitHub{bcolors.ENDC}")
    passed = False
    input()
    
if passed:
    if not os.path.exists("output"):
        os.makedirs("output")
    
    #Generate Images
    try:
        ppt_app = win32com.client.DispatchEx("PowerPoint.Application")       
        for file in os.listdir(directory):
            filename = os.fsdecode(file)
            if filename.endswith(".pptx"): 
                name = filename[:filename.rfind(".")]
                img = Image.new("RGB", (64,64),(255,255,255))
                img.save("output/" + name + ".png", "PNG")
                
                #Powerpoint stuff 
                export_slide(ppt_app,os.path.join(directory, filename),"output",name + ".png")
                names.append(str("output/"+name + ".png"))
                print(f"{bcolors.OKGREEN}Presentation converted successfully!{bcolors.ENDC}")
    finally:
        time.sleep(1)
        try:
            if close_powerpoint:
                ppt_app.Quit()
        except:
            pass
        
    #Open Images
    p1 = Image.open(names[0]).convert("RGBA")
    p2 = Image.open(names[1]).convert("RGBA")
    p3 = Image.open(names[2]).convert("RGBA")
    p4 = Image.open(names[3]).convert("RGBA")

    #Calculate the width and height 
    h = p1.height + p4.height
    w = p2.width + p3.width + p4.width 
    print("")   
    print(f"Final Image Height: {h}px")
    print(f"Final Image Width: {w}px")
    print("") 
    #boxes
    #(horni sirka, horni vyska)
    b2 = (0, p1.height)
    b4 = (p2.width, p1.height,)
    b3 = (p2.width+p4.width, p1.height)
    b1 = (p2.width-int(round(p4.width/4)), 0)

    #Pasting
    fimg = Image.new("RGBA", (w,h),(255,255,255,0))

    fimg.paste(p2,b2)
    fimg.paste(p4,b4)
    fimg.paste(p3,b3)
    fimg.paste(p1,b1)

    #Displaying and saving
    fimg.save("output/" + "final_merged" + ".png", "PNG")
    save = fimg.show()
    savei = input("Save the image? (y/n) ").lower()
    if savei == "y" or savei == "yes":
        fimg.save("output/" + "final_merged" + ".png", "PNG")
        print(f"{bcolors.OKGREEN}Saved successfully!{bcolors.ENDC}")
    else:
        print("Not saving")