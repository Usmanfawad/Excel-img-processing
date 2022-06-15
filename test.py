import excel2img
from PIL import Image
import os
import shutil
import subprocess


#Current directory constant
current_dir = os.getcwd()

#Get all xlsx files in the current work directory
ranges = ["D8:N17","P8:R17","T8:AB25"]

#Get all xlsx files in the current work directory
all_xlsx_files = []
for x in os.listdir():
    if x.endswith(".xlsx"):
        if not x[:2] == "~$":
            all_xlsx_files.append(x)
print(all_xlsx_files)
print("------ Beginning to process CSV files! ------\n")

for each_excel_file in all_xlsx_files:

    #closing any opened excel files
    subprocess.call([r'close_excels.bat'])

    print("\n> Processing " + each_excel_file)
    count_img = 1

    for each_range in ranges:
        try:
            print(each_range)
            excel2img.export_img(each_excel_file, f"{str(count_img)}.png", "Algemeen", f"{each_range}")
            print(f"Image {count_img} saved!")
            count_img += 1
        except Exception as e:
            print(e)
            

    print("I'm here")
    # Open images and store them in a list
    images = [Image.open(x) for x in ['2.png', '3.png']]
    total_width = 0
    max_height = 0
    # find the width and height of the final image
    for img in images:
        total_width += img.size[0]
        max_height = max(max_height, img.size[1])
    # create a new image with the appropriate height and width
    new_img = Image.new('RGB', (total_width+9, max_height))
    # Write the contents of the new image
    current_width = 0
    for img in images:
        new_img.paste(img, (current_width,0))
        current_width += img.size[0]
    # Save the image
    new_img.save('NewImage.jpg')

    #Close opened images
    for each_image in images:
        each_image.close()

    #Resizing the merged horizontal and separate verticle 1.png to bring uniformity in the image.

    basewidth = 1000
    img = Image.open('NewImage.jpg')
    wpercent = (basewidth / float(img.size[0]))
    hsize = int((float(img.size[1]) * float(wpercent)))
    img = img.resize((basewidth, hsize), Image.Resampling.LANCZOS)
    img.save('resized_image.jpg')
    img.close()
    img_2 = Image.open('1.png')
    wpercent = (basewidth / float(img_2.size[0]))
    hsize = int((float(img_2.size[1]) * float(wpercent)))
    img_2 = img_2.resize((basewidth, hsize), Image.Resampling.LANCZOS)
    img_2.save('resized_image_2.jpg')
    img_2.close()

    #vertically concatenating the pictures


    images_list = ['resized_image_2.jpg', 'resized_image.jpg']
    imgs = [Image.open(i) for i in images_list]

    min_img_width = min(i.width for i in imgs)

    total_height = 0
    for i, img in enumerate(imgs):
        # If the image is larger than the minimum width, resize it
        if img.width > min_img_width:
            imgs[i] = img.resize((min_img_width, int(img.height / img.width * min_img_width)), Image.ANTIALIAS)
        total_height += imgs[i].height

    # I have picked the mode of the first image to be generic. 
    # Now that we know the total height of all of the resized images, we know the height of our final image
    img_merge = Image.new(imgs[0].mode, (min_img_width, total_height))
    y = 0
    for img in imgs:
        img_merge.paste(img, (0, y))

        y += img.height

    # os.mkdir(each_excel_file[:-5]) 
    img_merge.save(f"excel_pngs\{each_excel_file[:-5]}.png")

    for each_img in imgs:
        each_img.close()

    #Deleting the unwanted files

    os.remove(os.path.join(current_dir,"1.png"))
    os.remove(os.path.join(current_dir,"2.png"))
    os.remove(os.path.join(current_dir,"3.png"))
    os.remove(os.path.join(current_dir,"NewImage.jpg"))
    os.remove(os.path.join(current_dir,"resized_image.jpg"))
    os.remove(os.path.join(current_dir,"resized_image_2.jpg"))
    

    # Moving current excel file to other folder
    curr_excel = current_dir + f"\{each_excel_file}"
    dest_excel = current_dir + f"\completed_excel\{each_excel_file}"
    shutil.move(curr_excel, dest_excel)


    print(">> Completed - " + each_excel_file)
    break