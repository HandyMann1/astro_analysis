import os
import tkinter as tk
from multiprocessing import Manager, Pool
from tkinter import filedialog

import cv2
import numpy as np
from openpyxl import Workbook

input_path = ""
output_path = ""


def load_images(image_folder):
    image_paths = [os.path.join(image_folder, f) for f in os.listdir(image_folder) if f.endswith('.jpg')]
    return image_paths


def analyze_image_chunk(image_chunk: tuple[cv2.Mat, int, str], stats_list):
    base_name = os.path.basename(image_chunk[2])
    new_name = os.path.splitext(base_name)[0]
    chunk_name = new_name + '_chunk_' + str(image_chunk[1])
    print(f"working on {chunk_name}")
    image = image_chunk[0]

    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    blurred = cv2.GaussianBlur(gray, (5, 5), 0)

    _, thresh = cv2.threshold(blurred, 100, 255, cv2.THRESH_BINARY)

    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    for contour in contours:
        area = cv2.contourArea(contour)
        if area < 2:
            continue

        mask = np.zeros(gray.shape, dtype=np.uint8)
        cv2.drawContours(mask, [contour], -1, 255, -1)

        mean_brightness = cv2.mean(gray, mask=mask)[0]

        (x, y), radius = cv2.minEnclosingCircle(contour)

        if area > 300:
            object_type = "Galaxy"
            color = (255, 0, 0)
        elif mean_brightness < 155:
            object_type = "Planet"
            color = (0, 255, 0)
        else:
            object_type = "Star"
            color = (0, 0, 255)

        stats_list.append({
            'Image': os.path.basename(image_chunk[2]),
            'chunk_number': image_chunk[1],
            'Coordinates': (x, y),
            'brightness': mean_brightness,
            'Object Type': object_type,
            'Area': area,
            'Radius': radius
        })
        cv2.drawContours(image, [contour], -1, color, thickness=2)

        M = cv2.moments(contour)
        if M["m00"] != 0:
            cX = int(M["m10"] / M["m00"])
            cY = int(M["m01"] / M["m00"])
            cv2.putText(image, object_type, (cX - 20, cY),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.5, color, 1)

        print(f"DONE working on {chunk_name}")
    cv2.imwrite(os.path.join(output_path, str(image_chunk[1]) + ".jpg"), image)


def save_to_excel(stats_list):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Astronomical Object Stats"

    sheet.append(
        ['Image', "part number", 'X Coordinate', 'Y Coordinate', 'brightness', 'Object Type', 'Area', 'Radius'])

    for stat in stats_list:
        x_coord, y_coord = stat['Coordinates']
        sheet.append([stat['Image'], stat['chunk_number'], x_coord, y_coord, stat['brightness'], stat['Object Type'],
                      stat['Area'],
                      stat['Radius']])

    output_file = os.path.join(output_path, "astronomical_object_stats.xlsx")

    try:
        workbook.save(output_file)
        print(f"Statistics saved to {output_file}")
    except Exception as e:
        print(f"Failed to save Excel file: {e}")


def split_image(image, part_size, image_path):
    image_height, image_width, _ = image.shape
    image_parts = []

    for offset_y in range(0, image_height, part_size):
        for offset_x in range(0, image_width, part_size):
            end_y = min(offset_y + part_size, image_height)
            end_x = min(offset_x + part_size, image_width)

            part = image[offset_y:end_y, offset_x:end_x]
            part_index = len(image_parts)

            image_parts.append((part, part_index, image_path))

    return image_parts


def analization():
    manager = Manager()
    stats_list = manager.list()
    image_parts = []
    image_paths = load_images(input_path)

    for image_path in image_paths:
        image: cv2.Mat = cv2.imread(image_path)
        segments: list[tuple[cv2.Mat, int, str]] = split_image(image, 500, image_path)
        for segment in segments:
            image_parts.append(segment)

    with Pool(processes=os.cpu_count()) as pool:
        pool.starmap(analyze_image_chunk, [(chunk, stats_list) for chunk in image_parts])
    print("donedonedone")

    save_to_excel(stats_list)


def select_input_directory():
    global input_path
    input_path = filedialog.askdirectory(
        initialdir="'D:\python_projects",
        title="Выберите файл"
    )
    input_file_name.config(text=input_path)


def select_output_directory():
    global output_path
    output_path = filedialog.askdirectory(
        initialdir="'D:\python_projects",
        title="Выберите файл"
    )
    output_file_name.config(text=output_path)


if __name__ == "__main__":
    root = tk.Tk()
    root.title("Анализ изображений")
    root.geometry("800x600")
    main_menu = tk.Menu(root)
    main_menu.add_cascade(label="Анализировать", command=analization)

    input_file_name_lbl = tk.Label(text="Имя файла ввода", font=("Helvetica", 16, "bold"))
    input_file_name_lbl.pack(anchor="w")

    input_file_name = tk.Label(text=input_path)
    input_file_name.pack(anchor="w")

    input_label_btn = tk.Button(root, text="Выберите ввод", command=select_input_directory)
    input_label_btn.pack(anchor="w")

    output_file_name_lbl = tk.Label(text="Имя файла вывода", font=("Helvetica", 16, "bold"))
    output_file_name_lbl.pack(anchor="w")

    output_file_name = tk.Label(text=output_path)
    output_file_name.pack(anchor="w")

    output_btn = tk.Button(text="Выберите вывод", command=select_output_directory)
    output_btn.pack(anchor="w")

    analyze_btn = tk.Button(text="Анализировать", command=analization)
    analyze_btn.pack(anchor="w")

    root.config(menu=main_menu)

    root.mainloop()
