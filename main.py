import win32com.client
import os
from PIL import Image

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
 
#Gemini helped with this powerpoint stuff   
def export_slide(pptx_path, output_folder,file):
    # Initialize PowerPoint
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    
    # Open the presentation
    abs_path = os.path.abspath(pptx_path)
    presentation = ppt_app.Presentations.Open(abs_path, WithWindow=False)
    
    # Export the first slide (Index starts at 1)
    slide = presentation.Slides(1)
    output_path = os.path.join(os.path.abspath(output_folder), file)
    
    # Export method (FileName, FilterName, Width, Height)
    slide.Export(output_path, "PNG")
    
    # Clean up
    presentation.Close()
    ppt_app.Quit()
    print(f"Exported to: {bcolors.OKBLUE}{output_path}{bcolors.ENDC}")


directory = os.fsencode(os.path.dirname(os.path.abspath(__file__)))
   
for file in os.listdir(directory):
    filename = os.fsdecode(file)
    if filename.endswith(".pptx"): 
        name = filename[:filename.rfind(".")]
        img = Image.new("RGB", (64,64),(255,255,255))
        img.save("output/" + name + ".png", "PNG")
        export_slide(filename,"output",name + ".png")
        print(f"{bcolors.OKGREEN}Presentation converted succesfully!{bcolors.ENDC}")


#Image merge
p1 = Image.open("output/ZIMČÍK_L_horni_2025.png").convert("RGBA")
p2 = Image.open("output/ZIMČÍK_L_leva_2025.png").convert("RGBA")
p3 = Image.open("output/ZIMČÍK_L_prava_2025.png").convert("RGBA")
p4 = Image.open("output/ZIMČÍK_L_stred_2025.png").convert("RGBA")


#Export everything
posterimg = [p1,p2,p3,p4]

#Calculate the width and height 
h = p1.height + p4.height
w = p2.width + p3.width + p4.width 
print("")   
print(f"Final Image Height: {h}")
print(f"Final Image Width: {w}")
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
    fimg.save("output/" + "final_mergedx" + ".png", "PNG")
    print(f"{bcolors.OKGREEN}Saved successfully!{bcolors.ENDC}")
else:
    print("Not saving")
