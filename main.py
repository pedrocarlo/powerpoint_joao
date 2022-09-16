import os
from tkinter import Tk, StringVar, OptionMenu, Grid, Frame, Button, Text, Label, X, LEFT, RIGHT, DISABLED
from tkinter.filedialog import askopenfilenames, askdirectory
from pptx import Presentation
from pptx.util import Inches
from templates import new_presentation
from PIL import Image, ImageTk

CWD = os.path.dirname(os.path.realpath(__file__))  # current working directory
root = Tk()
root.geometry("1000x700")

slide_frame = Frame(root)
initial_frame = Frame(root)
templates_frame = Frame(root)

slides = []  # contains the reference to


def main():
    create_initial_frame()
    root.mainloop()


def create_initial_frame():
    title = Text(initial_frame, height=1, width=15)
    title_label = Label(initial_frame, text="Put your presentation Title")

    new_prs_button = Button(master=initial_frame, text="New Presentation",
                            command=lambda: create_templates_frame(initial_frame, title))

    title_label.grid(row=0, column=0)
    title.grid(row=1, column=0)
    new_prs_button.grid(row=2, column=0)

    initial_frame.grid(sticky="")
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)


def create_templates_frame(frame, title_widget):
    global slides, template_names, template_files
    new_prs = new_presentation(title_widget.get("1.0", "end-1c"))
    slides = []

    frame.grid_forget()

    template_files = []
    while not template_files:
        template_dir = askdirectory(title='Select Template Directory', initialdir=CWD)
        # template_files = askopenfilenames(title='Select All Templates You Want To Use', initialdir=CWD,
        #                                   filetypes=[("PowerPoint Presentation", "*.pptx")])
        template_files = os.listdir(template_dir)
        for template_file in template_files:
            if not template_file.endswith(".pptx"):
                template_files = None
                break

    template_names = list(map(lambda x: os.path.basename(x), template_files))  # stores the PowerPoint templates names

    val = StringVar(templates_frame)
    val.set(template_names[0])

    menu = OptionMenu(templates_frame, val, *template_names)  # Option menu to select templates
    # menu.grid(row=0, column=0, sticky="NSEW")

    template_button = Button(templates_frame, text="Add Template Slide", command=lambda: add_slide(val.get()))

    templates_frame.grid(row=1, column=1, sticky="")
    for c in templates_frame.children:
        templates_frame.children[c].grid()
    root.rowconfigure((0, 1, 2), weight=1)
    root.columnconfigure((0, 1, 2, 3), weight=1)


def add_slide(template_name):
    global slides, template_names, template_files
    template_file = template_files[template_names.index(template_name)]  # getting the template file to add to queue
    slides.append(template_file)

    temp_frame = Frame(slide_frame)
    slide_button = Button(temp_frame, text=os.path.splitext(template_name)[0], command=lambda: image_viewer("sdf"))
    slide_button.grid(row=0, column=1)

    label_num = Label(temp_frame, text=str(len(slides)))
    label_num.grid(row=0, column=0)

    temp_frame.pack()
    slide_frame.grid(row=1, column=2, sticky="")


# code I got from geek for geeks for image viewer

def image_viewer(image_paths):
    templates_frame.grid_forget()
    slide_frame.grid_forget()
    image_paths = askopenfilenames(initialdir=CWD)
    image_index = 1

    # TODO need to do a check here for the number of images required for each template

    def add_more_images():
        return 0

    def new_image(direction: int):
        nonlocal label, button_forward, button_back, image_index, number_label
        label.grid_forget()

        if image_index + direction >= len(image_paths):
            button_forward = Button(button_frame, text="Forward", state=DISABLED)
            button_back = Button(button_frame, text="Back", command=lambda: new_image(-1))
        elif image_index + direction <= 1:
            button_forward = Button(button_frame, text="Forward", command=lambda: new_image(1))
            button_back = Button(button_frame, text="Back", state=DISABLED)
        else:
            button_forward = Button(button_frame, text="Forward", command=lambda: new_image(1))
            button_back = Button(button_frame, text="Back", command=lambda: new_image(-1))
        image_index += direction
        # This is for clearing the screen so that
        # our next image can pop up
        label = Label(label_frame, image=list_images[image_index - 1])
        # as the list starts from 0, so we are
        # subtracting one
        label.grid(row=1, column=0, columnspan=3)

        number_label = Label(label_frame, text=image_index)
        number_label.grid(row=0, column=0, columnspan=3, sticky="N")

        # Placing the button in new grid
        button_back.grid(row=5, column=0)
        # button_exit.grid(row=5, column=1)
        button_forward.grid(row=5, column=2)

    # Calling the Tk (The initial constructor of tkinter)
    image_frame = Frame(root)

    voltar = Button(image_frame, text="Voltar")
    voltar.grid(row=0, sticky="W")

    # List of the images so that we traverse the list
    list_images = []

    # Adding the images using the pillow module which
    # has a class ImageTk We can directly add the
    # photos in the tkinter folder or we have to
    # give a proper path for the images
    for path in image_paths:
        i = Image.open(path).resize((500, 500))
        list_images.append(ImageTk.PhotoImage(i))

    label_frame = Frame(image_frame, width=600, height=600)

    number_label = Label(label_frame, text=image_index)
    number_label.grid(row=0, column=0, columnspan=3, sticky="N")

    label_frame.grid(row=1, column=0, columnspan=3, sticky="N")
    label = Label(label_frame, image=list_images[0])

    # We have to show the box so this below line is needed
    label.grid(row=2, column=0, columnspan=3)

    button_frame = Frame(image_frame)
    button_frame.grid(row=1, column=0, columnspan=3, sticky="S")
    # We will have three button back ,forward and exit
    button_back = Button(button_frame, text="Back", command=lambda: new_image(-1),
                         state=DISABLED)

    # root.quit for closing the app
    # button_exit = Button(button_frame, text="Exit",
    #                      command=image_frame.quit)

    button_forward = Button(button_frame, text="Forward",
                            command=lambda: new_image(1))

    # grid function is for placing the buttons in the frame
    button_back.grid(row=5, column=0)
    # button_exit.grid(row=5, column=1)
    button_forward.grid(row=5, column=2)

    image_frame.grid()
    image_frame.rowconfigure(0, weight=3)
    image_frame.rowconfigure(1, weight=1)


if __name__ == '__main__':
    main()
