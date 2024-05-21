# This Python file uses the following encoding: utf-8
import sys

from PySide6.QtWidgets import QApplication, QWidget
import pandas as pd
import numpy as np

import openpyxl

# Important:
# You need to run the following command to generate the ui_form.py file
#     pyside6-uic form.ui -o ui_form.py, or
#     pyside2-uic form.ui -o ui_form.py
from ui_form import Ui_Widget

class Widget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.lb = Ui_Widget()
        self.ui = Ui_Widget()
        self.ui.setupUi(self)

        # Connect the button click event to the slot (function)
        self.ui.btn_read.clicked.connect(self.on_btn_read_clicked)
        self.ui.btn_analise.clicked.connect(self.on_btn_analise_clicked)
        self.ui.btn_help.clicked.connect(self.on_btn_help_clicked)

    def on_btn_read_clicked(self):
        text = "Reading.....!!!"
        self.ui.lb.setText(f'<font size="6">{text}</font>')
        # This function will be executed when the button is pressed
        print("Button Read pressed!")
        self.data = pd.read_excel('D:\wave_data.xlsx')

        # Extract wave height and direction columns
        self.wave_heights = self.data['Hs']
        self.directions = self.data['Dir']
        self.peak_period = self.data['Tp']

        # Check data range for wave heights and directions
        self.wave_height_range = (self.wave_heights.min(), self.wave_heights.max())
        self.direction_range = (self.directions.min(), self.directions.max())
        self.period_range = (self.peak_period.min(), self.peak_period.max())

        # Print and input wave height and direction ranges
        print("Wave Height Range:", self.wave_height_range)
        print("Direction Range:", self.direction_range)
        print("Period Range:", self.period_range)
        self.info_text = f"Wave Height Range: {self.wave_height_range}\nPeriod Range: {self.period_range}\nDirection Range: {self.direction_range}"
        self.ui.lb.setText(self.info_text)
        self.ui.btn_analise.setEnabled(True)
    def on_btn_analise_clicked(self):
        # Define variables with default values
        Hsmax = 0.0
        Hsmin = 0.0
        Tpmax = 0.0
        Tpmin = 0.0
        Dirmin = 0.0
        Dirmax = 0.0
        self.Hs_bin = 1
        self.Tp_bin = 1
        self.Dir_bin = 1

        try:
            # Retrieve values from the GUI text boxes and convert to float
            Hsmax = float(self.ui.txt_Hsmax.text())
            Hsmin = float(self.ui.txt_Hsmin.text())
            Tpmax = float(self.ui.txt_Tpmax.text())
            Tpmin = float(self.ui.txt_Tpmin.text())
            Dirmin = float(self.ui.txt_Dirmin.text())
            Dirmax = float(self.ui.txt_Dirmax.text())

            if (
                    not self.ui.txt_Hsbin.text()
                    or not self.ui.txt_Tpbin.text()
                    or not self.ui.txt_Dirbin.text()
            ):
                raise ValueError("Empty input field detected")

            self.Hs_bin = float(self.ui.txt_Hsbin.text())
            self.Tp_bin = float(self.ui.txt_Tpbin.text())
            self.Dir_bin = float(self.ui.txt_Dirbin.text())

        except ValueError as e:
            # Handle the case where the conversion to float fails (empty string or invalid input)
            print(f"Error: {e}")
        # This function will be executed when the button is pressed


        self.input_wave_height_range = (Hsmin,Hsmax)
        self.input_direction_range = (Dirmin,Dirmax)
        self.input_period_range = (Tpmin,Tpmax)

        # User input of bin size


        # Create pivot table with specified bin ranges
        bins_wave_heights = np.arange(self.input_wave_height_range[0], self.input_wave_height_range[1] + self.Hs_bin,
                                      self.Hs_bin)
        bins_directions = np.arange(self.input_direction_range[0], self.input_direction_range[1] + self.Dir_bin,
                                    self.Dir_bin)
        bins_period = np.arange(self.input_period_range[0], self.input_period_range[1] + self.Tp_bin,
                                self.Tp_bin)

        pivot_table_Hs = self.data.groupby([pd.cut(self.wave_heights, bins=bins_wave_heights),
                                            pd.cut(self.directions, bins=bins_directions)]).size().unstack(fill_value=0)

        pivot_table_Tp = self.data.groupby([pd.cut(self.wave_heights, bins=bins_wave_heights),
                                            pd.cut(self.directions, bins=bins_directions)])['Tp'].mean().unstack(
            fill_value=0)

        pivot_table_Pe = self.data.groupby([pd.cut(self.wave_heights, bins=bins_wave_heights),
                                            pd.cut(self.peak_period, bins=bins_period)]).size().unstack(fill_value=0)

        # Add a 'Total' column and row
        pivot_table_Hs['Total'] = pivot_table_Hs.sum(axis=1)
        pivot_table_Hs.loc['Total'] = pivot_table_Hs.sum()
        pivot_table_Tp['Mean'] = pivot_table_Tp.mean(axis=1)
        pivot_table_Tp.loc['Mean'] = pivot_table_Tp.mean()
        pivot_table_Pe['Total'] = pivot_table_Pe.sum(axis=1)
        pivot_table_Pe.loc['Total'] = pivot_table_Pe.sum()

        # Calculate percentage of grand total for each cell
        grand_total = pivot_table_Hs.loc['Total', 'Total']
        pivot_table_percent = (pivot_table_Hs / grand_total) * 100
        grand_total_Pe = pivot_table_Pe.loc['Total', 'Total']
        pivot_table_percent_Pe = (pivot_table_Pe / grand_total) * 100

        # Format the values with two decimal point
        pivot_table_percent = pivot_table_percent.round(2)
        pivot_table_Tp = pivot_table_Tp.round(1)
        pivot_table_percent_Pe = pivot_table_percent_Pe.round(2)
        self.ui.btn_analise.setEnabled(False)

        new_workbook_path = 'D:\wave_table.xlsx'

        # Create a Pandas Excel writer using ExcelWriter
        with pd.ExcelWriter(new_workbook_path, engine='openpyxl') as writer:
            # Write the pivot_table DataFrame to a new sheet named 'PivotTableSheet'
            pivot_table_percent.to_excel(writer, sheet_name='Hs_Dir_Occur', index=True)
            pivot_table_Tp.to_excel(writer, sheet_name='Hs_Dir_Tp_mean', index=True)
            pivot_table_percent_Pe.to_excel(writer, sheet_name='Hs_Tp_Occur', index=True)

        # Print a message indicating successful writing
        print("New workbook created at:", new_workbook_path)
        text = "New workbook created at : " + new_workbook_path
        self.ui.lb.setText(f'<font size="2">{text}</font>')

    def on_btn_help_clicked(self):

        text = "Input File - D:/wave_data.xlsx <br> Header Columns = Hs ,Tp , Dir <br> More - email: eranga.nilupul@gmail.com"
        self.ui.lb.setText(f'<font size="2">{text}</font>')


if __name__ == "__main__":
    app = QApplication(sys.argv)
    widget = Widget()
    widget.show()
    sys.exit(app.exec())

