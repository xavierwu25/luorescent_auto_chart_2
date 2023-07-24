import os
import pandas as pd
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog, Button
import openpyxl


def open_excel_file_1():
    root = Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        plot_xy_curve_from_excel(file_path)

def open_excel_file_2():
    root = Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        plot_xy_curve_from_excel(file_path)


def plot_xy_curve_from_excel(excel_file):
    df = pd.read_excel(excel_file)
    if 'C' in df.columns:


        x_label_3 = df.columns[0]
        y_label_3 = df.columns[1]


        x_data_3 = df[x_label_3]
        y_data_3 = df[y_label_3]


        plt.plot(x_data_3, y_data_3, linestyle='-', color='g', label=get_cell_value(excel_file, 'B1'))


        plt.fill_between(x_data_3, y_data_3, color='green', alpha=0.3)


        x_label_4 = df.columns[2]
        y_label_4 = df.columns[3]


        x_data_4 = df[x_label_4]
        y_data_4 = df[y_label_4]


        plt.plot(x_data_4, y_data_4, linestyle='-', color='red', label=get_cell_value(excel_file, 'D1'))

        plt.fill_between(x_data_4, y_data_4, color='red', alpha=0.3)


        x_label_5 = df.columns[4]
        y_label_5 = df.columns[5]


        x_data_5 = df[x_label_5]
        y_data_5 = df[y_label_5]


        plt.plot(x_data_5, y_data_5, linestyle='-', color='blue', label=get_cell_value(excel_file, 'F1'))


        plt.fill_between(x_data_5, y_data_5, color='blue', alpha=0.3)

    plt.xlabel('X軸')
    plt.ylabel('Y軸')
    plt.title(chart_title)
    plt.grid(True)
    plt.legend()
    plt.show()

def plot_xy_curve_from_excel(excel_file):
    df = pd.read_excel(excel_file)
    if 'C' not in df.columns:

        x_label_3 = df.columns[0]
        y_label_3 = df.columns[1]


        x_data_3 = df[x_label_3]
        y_data_3 = df[y_label_3]


        plt.plot(x_data_3, y_data_3, linestyle='-', color='g', label=get_cell_value(excel_file, 'B1'))


        plt.fill_between(x_data_3, y_data_3, color='green', alpha=0.3)


        x_label_5 = df.columns[4]
        y_label_5 = df.columns[5]


        x_data_5 = df[x_label_5]
        y_data_5 = df[y_label_5]


        plt.plot(x_data_5, y_data_5, linestyle='-', color='r', label=get_cell_value(excel_file, 'F1'))


        plt.fill_between(x_data_5, y_data_5, color='red', alpha=0.3)

    plt.xlabel('X軸')
    plt.ylabel('Y軸')
    plt.grid(True)
    plt.legend()
    plt.show()



def get_cell_value(excel_file, cell):
    wb = openpyxl.load_workbook(excel_file)
    ws = wb.active
    return ws[cell].value

root = Tk()
root.title("螢光物質與安鵬產品評估")

button1 = Button(root, text="選取螢光物質", command=open_excel_file_1)
button1.pack(pady=20)
button2 = Button(root, text="選取安鵬產品", command=open_excel_file_2)
button2.pack(pady=10)

window_width = 500
window_height = 400
root.geometry(f"{window_width}x{window_height}")

root.mainloop()