################################################################################
# Report issues at https://github.com/Krupicova12Kase/TK_Poster_Viewer
# Created by Krupicova12Kase, AKA Máťa or luki
# Version 2.0
# MIT license 
# Copyright (c) 2026 Krupicova12Kase
################################################################################

#settings
close_powerpoint = True #Should the program close powerpoint when it's done with generating? Not closing it may cause some problems
install_modules = False #Should the program install required packages? Packages are on line below
packages = ["pillow"]

#Imports
import os
from PIL import Image
import time
import subprocess
import aspose.slides as slides

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

#Function to export slides
def export_slide(file_name,output_name):
    with slides.Presentation(file_name) as presentation:
        slide = presentation.slides[0]
        with slide.get_image(1,1) as image:
            image.save(output_name, slides.ImageFormat.PNG)
            print("saving")

#Check if files are valid
directory = os.path.dirname(os.path.abspath(__file__))

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
        for file in os.listdir(directory):
            filename = os.fsdecode(file)
            if filename.endswith(".pptx"): 

                name = filename[:filename.rfind(".")]
                #img = Image.new("RGB", (64,64),(255,255,255))
                #img.save("output/" + name + ".png", "PNG")
                finaldir = os.path.join(directory, filename)
                #Powerpoint stuff 
                export_slide(finaldir,"output/"+ name + ".png")
                names.append(str("output/"+ name + ".png"))
                print(f"{bcolors.OKGREEN}Presentation converted successfully!{bcolors.ENDC}")
    except Exception as e:
        print(e)
        input()
    finally:
        time.sleep(1)
        
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