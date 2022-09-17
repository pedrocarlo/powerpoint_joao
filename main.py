import os
from tkinter import Tk, StringVar, OptionMenu, Frame, Button, Text, Label, DISABLED, Scrollbar
from tkinter.filedialog import askopenfilenames, askdirectory

from PIL import Image, ImageTk

from templates import new_presentation, Slide, save_presentation as save_pres

CWD = os.path.dirname(os.path.realpath(__file__))  # current working directory
root = Tk()
root.geometry("1100x700")

slide_frame = Frame(root)
scrollbar_slide_frame = Scrollbar(slide_frame, orient="vertical")  # TODO fix scrollbar
slide_button_frames = []

initial_frame = Frame(root)
templates_frame = Frame(root)

slides = {}  # contains the reference to slides


def main():
    create_initial_frame()
    root.mainloop()


def create_initial_frame():
    title = Text(initial_frame, height=1, width=15)
    title_label = Label(initial_frame, text="Write your presentation Title")

    new_prs_button = Button(master=initial_frame, text="New Presentation",
                            command=lambda: curry_check_title_pres(initial_frame, title.get('1.0', "end-1c")))

    title_label.grid(row=0, column=0)
    title.grid(row=1, column=0)
    new_prs_button.grid(row=2, column=0)

    initial_frame.grid(sticky="")
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)


def curry_check_title_pres(parent_frame, title):
    if title == "":
        blank_title = Label(parent_frame, text="Title Cannot be Blank", bg="red")
        blank_title.grid(row=3, column=0)
    else:
        create_templates_frame(parent_frame, title)


def create_templates_frame(frame, title):
    global slides, template_names, template_files
    slides = {}

    frame.grid_forget()

    template_files = []
    template_names = []
    while not template_files:
        template_dir = ""
        while template_dir == "":
            template_dir = askdirectory(title='Select Template Directory', initialdir=CWD)
        for template_file in os.listdir(template_dir):
            if not template_file.endswith(".pptx"):
                template_files = None
                break
        template_names = os.listdir(template_dir)  # stores the PowerPoint templates names
        template_files = list(map(lambda x: template_dir + "/" + x, template_names))

    val = StringVar(templates_frame)
    val.set(template_names[0])

    menu = OptionMenu(templates_frame, val, *template_names)  # Option menu to select templates
    # menu.grid(row=0, column=0, sticky="NSEW")

    template_button = Button(templates_frame, text="Add Template Slide",
                             command=lambda: add_new_slide(val.get()))

    save_button = Button(templates_frame, text="Save Presentation",
                         command=lambda: save_presentation(title))

    templates_frame.grid(row=1, column=1, sticky="")
    for c in templates_frame.children:
        templates_frame.children[c].grid()
    root.rowconfigure((0, 1, 2), weight=1)
    root.columnconfigure((0, 1, 2, 3), weight=1)

    save_button.grid(pady=40)


def save_presentation(title):
    new_prs, title = new_presentation(title)
    save_pres(new_prs, slides, title)


def add_new_slide(template_name):
    global slides, template_names, template_files
    template_file = template_files[template_names.index(template_name)]  # getting the template file to add to queue
    slide = Slide(template_file)
    add_slide(slide, template_name)


def add_existing_slide(slide: Slide, template_name: str):
    add_slide(slide, template_name)


def add_slide(slide: Slide, template_name: str):
    slides[len(slides.keys()) + 1] = slide

    temp_frame = create_slide_button_frame(template_name)
    slide_button_frames.append(temp_frame)

    slide_frame.grid(row=1, column=2, sticky="")
    # scrollbar_slide_frame.grid(row=0, column=10, sticky="W", rowspan=100)


def organize_slides():
    global slides
    keys = sorted(slides.keys())
    slides_to_sort = []
    for key in keys:
        slides_to_sort.append(slides[key])
    slides = {key + 1: slides_to_sort[key] for key in range(0, len(keys))}


def create_slide_button_frame(template_name):
    temp_frame = Frame(slide_frame)
    slide_num = len(slides.keys())
    slide_button = Button(temp_frame, text=os.path.splitext(template_name)[0],
                          command=lambda: image_viewer(slide_num))
    slide_button.grid(row=0, column=1)

    label_num = Label(temp_frame, text=str(len(slides.keys())))
    label_num.grid(row=0, column=0)

    delete_slide_button = Button(temp_frame, text="X", bg="red", command=lambda: delete_slide(slide_num))
    delete_slide_button.grid(row=0, column=2, sticky="W")

    temp_frame.grid()
    return temp_frame


def delete_slide(slide_num: int):
    global slides
    slides.pop(slide_num)
    organize_slides()
    for slide_button_frame in slide_button_frames:
        slide_button_frame.destroy()
    copy_slides = slides.copy()
    slides = {}
    for key, slide in copy_slides.items():
        add_existing_slide(slide, os.path.basename(slide.template_file))


# code I got from geek for geeks for image viewer
def image_viewer(slide_num):
    templates_frame.grid_forget()
    slide_frame.grid_forget()
    slide = slides[slide_num]
    if slide.images is None:
        image_paths = ""
        while image_paths == "":
            image_paths = list(askopenfilenames(initialdir=CWD, filetypes=[("Image Files", "*.png *.jpeg *.jpg")]))
        slide.update_images(image_paths)
    num_img, num_text = slide.num_images, slide.num_text
    if slide.texts is None:
        slide.texts = ["" for _ in range(num_text)]
    list_images_length = len(slide.images)
    image_index = 1
    list_images = []
    button_forward = None
    button_back = None
    number_label = None
    label = None

    # Calling the Tk (The initial constructor of tkinter)
    image_frame = Frame(root)
    label_frame = Frame(image_frame, width=600, height=600)
    button_frame = Frame(image_frame)

    def reselect_images():  # TODO
        return 0

    def next_image(direction: int):
        nonlocal label, button_forward, button_back, image_index, number_label, list_images
        label.grid_forget()
        if image_index + direction >= list_images_length:
            button_forward = Button(button_frame, text="Forward", state=DISABLED)
            button_back = Button(button_frame, text="Back", command=lambda: next_image(-1))
        elif image_index + direction <= 1:
            button_forward = Button(button_frame, text="Forward", command=lambda: next_image(1))
            button_back = Button(button_frame, text="Back", state=DISABLED)
        else:
            button_forward = Button(button_frame, text="Forward", command=lambda: next_image(1))
            button_back = Button(button_frame, text="Back", command=lambda: next_image(-1))
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
        button_forward.grid(row=5, column=2)

        # create labels with image and text count

    def order_images():
        nonlocal list_images
        for path in slide.images:
            i = Image.open(path).resize((500, 500))
            list_images.append(ImageTk.PhotoImage(i))
        return list_images

    def place_image():
        number_label = Label(label_frame, text=image_index)
        number_label.grid(row=0, column=0, columnspan=3, sticky="N")

        label = Label(label_frame, image=list_images[0])

        # We have to show the box so this below line is needed
        label.grid(row=2, column=0, columnspan=3)
        label_frame.grid(row=1, column=0, columnspan=3, sticky="N")
        return label, number_label

    def create_menus_orders():
        menu_frame = Frame(image_frame)

        image_choice = StringVar(menu_frame)
        image_choice.set("1")
        if list_images_length >= num_img:
            image_tot = num_img
        else:
            image_tot = list_images_length
        image_menu = OptionMenu(menu_frame, image_choice, *range(1, image_tot + 1))
        image_button = Button(menu_frame, text="Change order of Displayed image to selected number order",
                              command=lambda: change_image_order(int(image_choice.get())))

        text_choice = StringVar(menu_frame)
        text_choice.set("1")
        text_menu = OptionMenu(menu_frame, text_choice, *range(1, num_text + 1))
        text_box = Text(menu_frame, width=30, height=2)
        text_button = Button(menu_frame, text="Set Text for selected text box",
                             command=lambda: set_text(text_box.get("1.0", "end-1c"),
                                                      int(text_choice.get()) - 1))

        menu_frame.grid(row=2, column=0, sticky="nsew")

        image_menu.grid(row=0, column=0)
        image_button.grid(row=0, column=1, columnspan=3)
        text_menu.grid(row=1, column=0)
        text_button.grid(row=1, column=1, columnspan=3)
        text_box.grid(row=1, column=5)

    def set_text(text, text_index):
        slide.texts[text_index] = text

    def change_image_order(new_index: int):
        nonlocal label, number_label, image_index, list_images
        if new_index == image_index:
            return
        list_images.insert(new_index - 1, list_images.pop(image_index - 1))
        slide.images.insert(new_index - 1, slide.images.pop(image_index - 1))
        image_index = 1
        for c in label_frame.children:
            label_frame.children[c].grid_forget()
        initial_state()

    def initial_state():
        nonlocal list_images, label, number_label, button_forward, button_back
        voltar = Button(image_frame, text="Voltar",
                        command=lambda: return_to_template_frame(image_frame, slide))
        voltar.grid(row=0, column=0, sticky="W")

        # List of the images so that we traverse the list
        list_images = order_images()

        label, number_label = place_image()  # puts label and number label for image

        # We will have three button back ,forward and exit
        button_back = Button(button_frame, text="Back", command=lambda: next_image(-1),
                             state=DISABLED)

        if list_images_length <= 1:  # accounting for the exception where you only have only have 1 picture
            button_forward = Button(button_frame, text="Forward", state=DISABLED)
        else:
            button_forward = Button(button_frame, text="Forward",
                                    command=lambda: next_image(1))

        # grid function is for placing the buttons in the frame
        button_back.grid(row=5, column=0)
        button_forward.grid(row=5, column=2)

        image_frame.grid()
        image_frame.rowconfigure(0, weight=3)
        image_frame.rowconfigure(1, weight=1)
        image_frame.rowconfigure(2, weight=2)

        label_frame.grid(row=1, column=0, columnspan=3, sticky="N")
        button_frame.grid(row=1, column=0, columnspan=3, sticky="S")

        image_count_label = Label(image_frame, text="Amount of Images on template: " + str(num_img))
        text_count_label = Label(image_frame, text="Amount of Text Boxes on template: " + str(num_text))
        image_count_label.grid(row=0, column=1, sticky="W")
        text_count_label.grid(row=0, column=2, sticky="W")

        create_menus_orders()

    initial_state()


def return_to_template_frame(image_frame, slide):
    slide.update_images(slide.get_images())
    image_frame.grid_forget()
    templates_frame.grid(row=1, column=1, sticky="")
    slide_frame.grid(row=1, column=2, sticky="")


if __name__ == '__main__':
    main()
