import sys
import tkinter as tk # To create UI
from tkinter import filedialog
import pandas as pd # To Read exl file
import matplotlib.pyplot as plt   # Draw graphs
import numpy as np
import xlwt
from xlwt import Workbook

from PyQt5.QtWidgets import QMainWindow, QApplication

from mainUI import *

# Workbook is created to save data
wb = Workbook()
# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Sheet 1', cell_overwrite_ok=True)
# Specifying style
styleBold = xlwt.easyxf('font: bold 1')


class MyForm(QMainWindow):

    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.ui.BrowseBtnSelection.clicked.connect(self.getExcelSelection)
        #shading
        self.ui.BrowseBtnShading_1.clicked.connect(self.getExcelShading_1)
        self.ui.BrowseBtnShading_2.clicked.connect(self.getExcelShading_2)
        self.ui.BrowseBtnShading_3.clicked.connect(self.getExcelShading_3)
        #Hotspot
        self.ui.BrowseBtnHotspot_1.clicked.connect(self.getExcelHotspot_1)
        self.ui.BrowseBtnHotspot_2.clicked.connect(self.getExcelHotspot_2)
        self.ui.BrowseBtnHotspot_3.clicked.connect(self.getExcelHotspot_3)

        self.show()

    def getExcelHotspot_1(self):
        global df_Hotspot_1

        import_file_path = filedialog.askopenfilename(title = "Select file",filetypes = (("Excel files","*.xlsx"),("all files","*.*")))

        df_Hotspot_1 = pd.read_excel(import_file_path)
        print(df_Hotspot_1)
        self.ui.hotspotFilePath_1.setText(import_file_path)

        self.readDataHotspot_1()

    def getExcelHotspot_2(self):
        global df_Hotspot_2

        import_file_path = filedialog.askopenfilename()
        df_Hotspot_2 = pd.read_excel(import_file_path)
        print(df_Hotspot_2)
        self.ui.hotspotFilePath_2.setText(import_file_path)
        self.readDataHotspot_2()

    def getExcelHotspot_3(self):
        global df_Hotspot_3

        import_file_path = filedialog.askopenfilename()
        df_Hotspot_3 = pd.read_excel(import_file_path)
        print(df_Hotspot_3)
        self.ui.hotspotFilePath_3.setText(import_file_path)
        self.readDataHotspot_3()

    def getExcelShading_1(self):
        global df_Shading_1

        import_file_path = filedialog.askopenfilename()
        df_Shading_1 = pd.read_excel(import_file_path)
        print(df_Shading_1)
        self.ui.shadingFilePath_1.setText(import_file_path)
        self.readDataShading_1()

    def getExcelShading_2(self):
        global df_Shading_2

        import_file_path = filedialog.askopenfilename()
        df_Shading_2 = pd.read_excel(import_file_path)
        print(df_Shading_2)
        self.ui.shadingFilePath_2.setText(import_file_path)
        self.readDataShading_2()


    def getExcelShading_3(self):
        global df_Shading_3

        import_file_path = filedialog.askopenfilename()
        df_Shading_3 = pd.read_excel(import_file_path)
        print(df_Shading_3)
        self.ui.shadingFilePath_3.setText(import_file_path)
        self.readDataShading_3()

    def getExcelSelection(self):
        global df_Selection

        import_file_path = filedialog.askopenfilename()
        df_Selection = pd.read_excel(import_file_path)
        print(df_Selection)
        self.ui.selectionFilePath.setText(import_file_path)
        self.readDataSelection()
        self.showChart()

    def showChart(self):
        df_Selection_1_chart = df_Selection[1:]
        df = df_Selection_1_chart

        # Saving the first 2 column data
        new_columns = ["X_axis", "Y_axis"]
        df_part = df.iloc[:, 0:2]
        df_part.columns = new_columns
        df1 = pd.DataFrame(df_part, columns=['Cell'] + new_columns)
        df1['Cell'] = df.columns[0]
        # print(df1)

        # Creating long range data from wide rangs so that now data is 13800 (230*60) by 3 instead of 230 by 120 data to be able to generate graphs
        for colnum in range(2, len(df.columns) - 1, 2):
            part1 = df1.iloc[:, 0:3]
            part2 = df.iloc[:, colnum:(colnum + 2)]
            part2.columns = new_columns
            part2['Cell'] = df.columns[colnum]
            # Re-order columns
            cols = part2.columns.tolist()
            cols = cols[-1:] + cols[:-1]
            part2 = part2[cols]
            # Merge the data from cols
            df1 = pd.concat([part1, part2], ignore_index=True)

        # print(df1)

        # Create Graph for all the Cells
        fig, ax = plt.subplots(figsize=(6, 7))
        df1.groupby('Cell').plot(kind='line', x='Y_axis', y='X_axis', ax=ax)
        # naming the x axis
        plt.xlabel('Voltage')
        # naming the y axis
        plt.ylabel('Current')
        # giving a title to my graph
        plt.title('Cell Selection!')
        # Remove legends
        ax.get_legend().remove()
        # Default limits
        # plt.ylim(0, 10)
        # plt.xlim(0, 35)
        # setting x and y-axis limits
        plt.ylim(0, 2)
        plt.xlim(20, 30)
        # Show the graph
        plt.show()


    def readDataSelection(self):
        a = [100] * 200
        b = [99] * 200
        preventer = 0
        store1 = [99] * 200
        store2 = [99] * 200
        from matplotlib import pyplot as pl
        sum = 0
        df_Selection.shape
        for i in range(0, 120, 2):
            dummy = df_Selection.iloc[0, i]
            #print(dummy)
            store1[i] = dummy
            first = df_Selection.iloc[1, i]

            second = df_Selection.iloc[1, i + 1]
            a[i] = second / first

            if (a[i] == 0):
                a[i] = 99
            #print(a[i])
            store2[i] = a[i]
        pos = store2.index(min(store2))
        print(store1[pos])
        cell_1 = store1[pos]
        print(store2[pos])

        store2[pos] = 99
        pos = store2.index(min(store2))
        print(store1[pos])
        cell_2 = store1[pos]
        print(store2[pos])
        store2[pos] = 99
        pos = store2.index(min(store2))
        print(store1[pos])
        cell_3 = store1[pos]
        print(store2[pos])
        #setting values to display
        self.ui.cell_1_val.setText(cell_1[4:])
        self.ui.cell_2_val.setText(cell_2[4:])
        self.ui.cell_3_val.setText(cell_3[4:])

        # SELECTION saving data
        sheet1.write(1, 3, 'THREE  SELECTED  CELLS', styleBold)
        sheet1.write(3, 3, 'LOW RSHUNT CELL 1', styleBold)
        sheet1.write(4, 3, 'LOW RSHUNT CELL 2', styleBold)
        sheet1.write(5, 3, 'LOW RSHUNT CELL 3', styleBold)
        # data
        sheet1.write(3, 4, cell_1[4:])
        sheet1.write(4, 4, cell_2[4:])
        sheet1.write(5, 4, cell_3[4:])

        wb.save('HotSpot_Test_Results.xls')


    def readDataShading_1(self):
        a = [99] * 100
        b = [99] * 3
        counter = 0
        from matplotlib import pyplot as pl
        sum = 0
        imp = 8.7
        n = 0
        errorprevent = 0

        df_Shading_1.shape
        for i in range(0, 6, 2):

            if (errorprevent != 0):
                sum = sum / n
                a[counter] = abs(sum - imp)
                b[counter] = min(a)
                counter = counter + 1

                print(sum)
                sum = 0
                n = 0

            for j in range(1, 230):
                first = df_Shading_1.iloc[j, i]
                if (imp - 3 < first < imp + 3):
                    sum = sum + first
                    n = n + 1
                errorprevent = 99
        print(b)
        if (b.index(min(b)) == 0):
            self.ui.cell_4_val.setText("0%")
            sheet1.write(3, 8, '0%')
           # print("0%")
        elif (b.index(min(b)) == 1):
            self.ui.cell_4_val.setText("12.5%")
            sheet1.write(3, 8, '12.5%')
            #print("12.5%")
        elif (b.index(min(b)) == 2):
            self.ui.cell_4_val.setText("25%")
            sheet1.write(3, 8, '25%')
            #print("25%")
        #Shading  #setting values to display
        # Shading
        sheet1.write(1, 7, 'SHADING RESULTS', styleBold)
        sheet1.write(3, 7, 'CELL 1', styleBold)
        sheet1.write(4, 7, 'CELL 2', styleBold)
        sheet1.write(5, 7, 'CELL 3', styleBold)
        # data
        wb.save('HotSpot_Test_Results.xls')

    def readDataShading_2(self):
        a = [99] * 100
        b = [99] * 3
        counter = 0
        from matplotlib import pyplot as pl
        sum = 0
        imp = 8.7
        n = 0
        errorprevent = 0

        df_Shading_2.shape
        for i in range(0, 6, 2):

            if (errorprevent != 0):
                sum = sum / n
                a[counter] = abs(sum - imp)
                b[counter] = min(a)
                counter = counter + 1

                print(sum)
                sum = 0
                n = 0

            for j in range(1, 230):
                first = df_Shading_2.iloc[j, i]
                if (imp - 3 < first < imp + 3):
                    sum = sum + first
                    n = n + 1
                errorprevent = 99
        print(b)
        if (b.index(min(b)) == 0):
            self.ui.cell_5_val.setText("0%")
            sheet1.write(4, 8, '0%')
           # print("0%")
        elif (b.index(min(b)) == 1):
            self.ui.cell_5_val.setText("25%")
            sheet1.write(4, 8, '25%')
            #print("12.5%")
        elif (b.index(min(b)) == 2):
            self.ui.cell_5_val.setText("25%")
            sheet1.write(4, 8, '25%')
            #print("25%")
        #Shading  #setting values to display
        # Shading
        sheet1.write(1, 7, 'SHADING RESULTS', styleBold)
        sheet1.write(3, 7, 'CELL 1', styleBold)
        sheet1.write(4, 7, 'CELL 2', styleBold)
        sheet1.write(5, 7, 'CELL 3', styleBold)
        # data
        wb.save('HotSpot_Test_Results.xls')

    def readDataShading_3(self):
        a = [99] * 100
        b = [99] * 3
        counter = 0
        from matplotlib import pyplot as pl
        sum = 0
        imp = 8.7
        n = 0
        errorprevent = 0

        df_Shading_3.shape
        for i in range(0, 6, 2):

            if (errorprevent != 0):
                sum = sum / n
                a[counter] = abs(sum - imp)
                b[counter] = min(a)
                counter = counter + 1

                print(sum)
                sum = 0
                n = 0

            for j in range(1, 230):
                first = df_Shading_3.iloc[j, i]
                if (imp - 3 < first < imp + 3):
                    sum = sum + first
                    n = n + 1
                errorprevent = 99
        print(b)
        if (b.index(min(b)) == 0):
            self.ui.cell_6_val.setText("0%")
            sheet1.write(5, 8, '0%')
           # print("0%")
        elif (b.index(min(b)) == 1):
            self.ui.cell_6_val.setText("12.5%")
            sheet1.write(5, 8, '12.5%')
            #print("12.5%")
        elif (b.index(min(b)) == 2):
            self.ui.cell_6_val.setText("25%")
            sheet1.write(5, 8, '25%')
            #print("25%")
        #Shading  #setting values to display
        # Shading
        sheet1.write(1, 7, 'SHADING RESULTS', styleBold)
        sheet1.write(3, 7, 'CELL 1', styleBold)
        sheet1.write(4, 7, 'CELL 2', styleBold)
        sheet1.write(5, 7, 'CELL 3', styleBold)
        # data
        wb.save('HotSpot_Test_Results.xls')



    def readDataHotspot_1(self):

        a = [99] * 100
        b = [99] * 3
        arrmin = [100] * 100
        arrmax = [0] * 100
        arravg = [100.0] * 100
        counter = 0
        from matplotlib import pyplot as pl

        imp = 8.7
        n = 0
        errorprevent = 0

        c = df_Hotspot_1.iloc[16:, 2]
        arrmin[0] = min(c)
        arrmax[0] = max(c)
        arravg[0] = np.mean(c)

        c = df_Hotspot_1.iloc[16:, 5]

        arrmin[1] = min(c)
        arrmax[1] = max(c)
        arravg[1] = np.mean(c)
        c = df_Hotspot_1.iloc[16:, 7]
        arrmin[2] = min(c)
        arrmax[2] = max(c)

        arravg[2] = np.mean(c)
        c = df_Hotspot_1.iloc[16:, 8]

        arrmin[3] = min(c)
        arrmax[3] = max(c)
        arravg[3] = np.mean(c)
        c = df_Hotspot_1.iloc[16:, 9]
        arrmin[4] = min(c)
        arrmax[4] = max(c)

        arravg[4] = np.mean(c)
        c = df_Hotspot_1.iloc[16:, 10]
        arrmin[5] = min(c)
        arrmax[5] = max(c)

        arravg[5] = np.mean(c)
        c = df_Hotspot_1.iloc[16:, 11]
        arrmin[6] = min(c)
        arrmax[6] = max(c)

        arravg[6] = np.mean(c)
        c = df_Hotspot_1.iloc[16:, 12]
        arrmin[7] = min(c)
        arrmax[7] = max(c)

        arravg[7] = np.mean(c)

        self.ui.cell_HS_11.setText(str(arrmax[0]))
        self.ui.cell_HS_12.setText(str(arrmin[0]))
        self.ui.cell_HS_13.setText(str(arravg[0]))

        self.ui.cell_HS_21.setText(str(arrmax[1]))
        self.ui.cell_HS_22.setText(str(arrmin[1]))
        self.ui.cell_HS_23.setText(str(arravg[1]))

        self.ui.cell_HS_31.setText(str(arrmax[2]))
        self.ui.cell_HS_32.setText(str(arrmin[2]))
        self.ui.cell_HS_33.setText(str(arravg[2]))

        self.ui.cell_HS_41.setText(str(arrmax[3]))
        self.ui.cell_HS_42.setText(str(arrmin[3]))
        self.ui.cell_HS_43.setText(str(arravg[3]))

        self.ui.cell_HS_51.setText(str(arrmax[4]))
        self.ui.cell_HS_52.setText(str(arrmin[4]))
        self.ui.cell_HS_53.setText(str(arravg[4]))

        self.ui.cell_HS_61.setText(str(arrmax[5]))
        self.ui.cell_HS_62.setText(str(arrmin[5]))
        self.ui.cell_HS_63.setText(str(arravg[5]))

        self.ui.cell_HS_71.setText(str(arrmax[6]))
        self.ui.cell_HS_72.setText(str(arrmin[6]))
        self.ui.cell_HS_73.setText(str(arravg[6]))

        self.ui.cell_HS_81.setText(str(arrmax[7]))
        self.ui.cell_HS_82.setText(str(arrmin[7]))
        self.ui.cell_HS_83.setText(str(arravg[7]))

        sheet1.write(8, 3, 'HOT-SPOT RESULTS CELL 1', styleBold)
        sheet1.write(11, 3, 'MAX', styleBold)
        sheet1.write(11, 4, 'MIN', styleBold)
        sheet1.write(11, 5, 'AVG', styleBold)
        sheet1.write(12, 1, 'T amb', styleBold)
        sheet1.write(13, 1, 'Ref. Cell Voltage', styleBold)
        sheet1.write(14, 1, 'T Cell 1', styleBold)
        sheet1.write(15, 1, 'T Cell 2', styleBold)
        sheet1.write(16, 1, 'T module', styleBold)
        sheet1.write(17, 1, 'Calc. Irradiance', styleBold)
        sheet1.write(18, 1, 'Calc. T rise,cell', styleBold)
        sheet1.write(19, 1, 'Calc. T rise,cel', styleBold)

        sheet1.write(12, 3, arrmax[0])
        sheet1.write(12, 4, arrmin[0])
        sheet1.write(12, 5, arravg[0])

        sheet1.write(13, 3, arrmax[1])
        sheet1.write(13, 4, arrmin[1])
        sheet1.write(13, 5, arravg[1])

        sheet1.write(14, 3, arrmax[2])
        sheet1.write(14, 4, arrmin[2])
        sheet1.write(14, 5, arravg[2])

        sheet1.write(15, 3, arrmax[3])
        sheet1.write(15, 4, arrmin[3])
        sheet1.write(15, 5, arravg[3])

        sheet1.write(16, 3, arrmax[4])
        sheet1.write(16, 4, arrmin[4])
        sheet1.write(16, 5, arravg[4])

        sheet1.write(17, 3, arrmax[5])
        sheet1.write(17, 4, arrmin[5])
        sheet1.write(17, 5, arravg[5])

        sheet1.write(18, 3, arrmax[6])
        sheet1.write(18, 4, arrmin[6])
        sheet1.write(18, 5, arravg[6])

        sheet1.write(19, 3, arrmax[7])
        sheet1.write(19, 4, arrmin[7])
        sheet1.write(19, 5, arravg[7])

        wb.save('HotSpot_Test_Results.xls')

    def readDataHotspot_2(self):

        a = [99] * 100
        b = [99] * 3
        arrmin = [100] * 100
        arrmax = [0] * 100
        arravg = [100.0] * 100
        counter = 0
        from matplotlib import pyplot as pl

        imp = 8.7
        n = 0
        errorprevent = 0

        c = df_Hotspot_2.iloc[16:, 2]
        arrmin[0] = min(c)
        arrmax[0] = max(c)
        arravg[0] = np.mean(c)

        c = df_Hotspot_2.iloc[16:, 5]

        arrmin[1] = min(c)
        arrmax[1] = max(c)
        arravg[1] = np.mean(c)
        c = df_Hotspot_2.iloc[16:, 7]
        arrmin[2] = min(c)
        arrmax[2] = max(c)

        arravg[2] = np.mean(c)
        c = df_Hotspot_2.iloc[16:, 8]

        arrmin[3] = min(c)
        arrmax[3] = max(c)
        arravg[3] = np.mean(c)
        c = df_Hotspot_2.iloc[16:, 9]
        arrmin[4] = min(c)
        arrmax[4] = max(c)

        arravg[4] = np.mean(c)
        c = df_Hotspot_2.iloc[16:, 10]
        arrmin[5] = min(c)
        arrmax[5] = max(c)

        arravg[5] = np.mean(c)
        c = df_Hotspot_2.iloc[16:, 11]
        arrmin[6] = min(c)
        arrmax[6] = max(c)

        arravg[6] = np.mean(c)
        c = df_Hotspot_2.iloc[16:, 12]
        arrmin[7] = min(c)
        arrmax[7] = max(c)

        arravg[7] = np.mean(c)

        self.ui.cell_HS_11.setText(str(arrmax[0]))
        self.ui.cell_HS_12.setText(str(arrmin[0]))
        self.ui.cell_HS_13.setText(str(arravg[0]))

        self.ui.cell_HS_21.setText(str(arrmax[1]))
        self.ui.cell_HS_22.setText(str(arrmin[1]))
        self.ui.cell_HS_23.setText(str(arravg[1]))

        self.ui.cell_HS_31.setText(str(arrmax[2]))
        self.ui.cell_HS_32.setText(str(arrmin[2]))
        self.ui.cell_HS_33.setText(str(arravg[2]))

        self.ui.cell_HS_41.setText(str(arrmax[3]))
        self.ui.cell_HS_42.setText(str(arrmin[3]))
        self.ui.cell_HS_43.setText(str(arravg[3]))

        self.ui.cell_HS_51.setText(str(arrmax[4]))
        self.ui.cell_HS_52.setText(str(arrmin[4]))
        self.ui.cell_HS_53.setText(str(arravg[4]))

        self.ui.cell_HS_61.setText(str(arrmax[5]))
        self.ui.cell_HS_62.setText(str(arrmin[5]))
        self.ui.cell_HS_63.setText(str(arravg[5]))

        self.ui.cell_HS_71.setText(str(arrmax[6]))
        self.ui.cell_HS_72.setText(str(arrmin[6]))
        self.ui.cell_HS_73.setText(str(arravg[6]))

        self.ui.cell_HS_81.setText(str(arrmax[7]))
        self.ui.cell_HS_82.setText(str(arrmin[7]))
        self.ui.cell_HS_83.setText(str(arravg[7]))

        # Hotspot 2
        sheet1.write(8, 9, 'HOT-SPOT RESULTS CELL 2', styleBold)
        sheet1.write(11, 9, 'MAX', styleBold)
        sheet1.write(11, 10, 'MIN', styleBold)
        sheet1.write(11, 11, 'AVG', styleBold)
        sheet1.write(12, 7, 'T amb', styleBold)
        sheet1.write(13, 7, 'Ref. Cell Voltage', styleBold)
        sheet1.write(14, 7, 'T Cell 1', styleBold)
        sheet1.write(15, 7, 'T Cell 2', styleBold)
        sheet1.write(16, 7, 'T module', styleBold)
        sheet1.write(17, 7, 'Calc. Irradiance', styleBold)
        sheet1.write(18, 7, 'Calc. T rise,cell', styleBold)
        sheet1.write(19, 7, 'Calc. T rise,cel', styleBold)

        sheet1.write(12, 9, arrmax[0])
        sheet1.write(12, 10, arrmin[0])
        sheet1.write(12, 11, arravg[0])

        sheet1.write(13, 9, arrmax[1])
        sheet1.write(13, 10, arrmin[1])
        sheet1.write(13, 11, arravg[1])

        sheet1.write(14, 9, arrmax[2])
        sheet1.write(14, 10, arrmin[2])
        sheet1.write(14, 11, arravg[2])

        sheet1.write(15, 9, arrmax[3])
        sheet1.write(15, 10, arrmin[3])
        sheet1.write(15, 11, arravg[3])

        sheet1.write(16, 9, arrmax[4])
        sheet1.write(16, 10, arrmin[4])
        sheet1.write(16, 11, arravg[4])

        sheet1.write(17, 9, arrmax[5])
        sheet1.write(17, 10, arrmin[5])
        sheet1.write(17, 11, arravg[5])

        sheet1.write(18, 9, arrmax[6])
        sheet1.write(18, 10, arrmin[6])
        sheet1.write(18, 11, arravg[6])

        sheet1.write(19, 9, arrmax[7])
        sheet1.write(19, 10, arrmin[7])
        sheet1.write(19, 11, arravg[7])

        wb.save('HotSpot_Test_Results.xls')

    def readDataHotspot_3(self):

        a = [99] * 100
        b = [99] * 3
        arrmin = [100] * 100
        arrmax = [0] * 100
        arravg = [100.0] * 100
        counter = 0
        from matplotlib import pyplot as pl

        imp = 8.7
        n = 0
        errorprevent = 0

        c = df_Hotspot_3.iloc[16:, 2]
        arrmin[0] = min(c)
        arrmax[0] = max(c)
        arravg[0] = np.mean(c)

        c = df_Hotspot_3.iloc[16:, 5]

        arrmin[1] = min(c)
        arrmax[1] = max(c)
        arravg[1] = np.mean(c)
        c = df_Hotspot_3.iloc[16:, 7]
        arrmin[2] = min(c)
        arrmax[2] = max(c)

        arravg[2] = np.mean(c)
        c = df_Hotspot_3.iloc[16:, 8]

        arrmin[3] = min(c)
        arrmax[3] = max(c)
        arravg[3] = np.mean(c)
        c = df_Hotspot_3.iloc[16:, 9]
        arrmin[4] = min(c)
        arrmax[4] = max(c)

        arravg[4] = np.mean(c)
        c = df_Hotspot_3.iloc[16:, 10]
        arrmin[5] = min(c)
        arrmax[5] = max(c)

        arravg[5] = np.mean(c)
        c = df_Hotspot_3.iloc[16:, 11]
        arrmin[6] = min(c)
        arrmax[6] = max(c)

        arravg[6] = np.mean(c)
        c = df_Hotspot_3.iloc[16:, 12]
        arrmin[7] = min(c)
        arrmax[7] = max(c)

        arravg[7] = np.mean(c)

        self.ui.cell_HS_11.setText(str(arrmax[0]))
        self.ui.cell_HS_12.setText(str(arrmin[0]))
        self.ui.cell_HS_13.setText(str(arravg[0]))

        self.ui.cell_HS_21.setText(str(arrmax[1]))
        self.ui.cell_HS_22.setText(str(arrmin[1]))
        self.ui.cell_HS_23.setText(str(arravg[1]))

        self.ui.cell_HS_31.setText(str(arrmax[2]))
        self.ui.cell_HS_32.setText(str(arrmin[2]))
        self.ui.cell_HS_33.setText(str(arravg[2]))

        self.ui.cell_HS_41.setText(str(arrmax[3]))
        self.ui.cell_HS_42.setText(str(arrmin[3]))
        self.ui.cell_HS_43.setText(str(arravg[3]))

        self.ui.cell_HS_51.setText(str(arrmax[4]))
        self.ui.cell_HS_52.setText(str(arrmin[4]))
        self.ui.cell_HS_53.setText(str(arravg[4]))

        self.ui.cell_HS_61.setText(str(arrmax[5]))
        self.ui.cell_HS_62.setText(str(arrmin[5]))
        self.ui.cell_HS_63.setText(str(arravg[5]))

        self.ui.cell_HS_71.setText(str(arrmax[6]))
        self.ui.cell_HS_72.setText(str(arrmin[6]))
        self.ui.cell_HS_73.setText(str(arravg[6]))

        self.ui.cell_HS_81.setText(str(arrmax[7]))
        self.ui.cell_HS_82.setText(str(arrmin[7]))
        self.ui.cell_HS_83.setText(str(arravg[7]))

        # Hotspot 3
        sheet1.write(8, 15, 'HOT-SPOT RESULTS CELL 3', styleBold)
        sheet1.write(11, 15, 'MAX', styleBold)
        sheet1.write(11, 16, 'MIN', styleBold)
        sheet1.write(11, 17, 'AVG', styleBold)
        sheet1.write(12, 13, 'T amb', styleBold)
        sheet1.write(13, 13, 'Ref. Cell Voltage', styleBold)
        sheet1.write(14, 13, 'T Cell 1', styleBold)
        sheet1.write(15, 13, 'T Cell 2', styleBold)
        sheet1.write(16, 13, 'T module', styleBold)
        sheet1.write(17, 13, 'Calc. Irradiance', styleBold)
        sheet1.write(18, 13, 'Calc. T rise,cell', styleBold)
        sheet1.write(19, 13, 'Calc. T rise,cel', styleBold)

        sheet1.write(12, 15, arrmax[0])
        sheet1.write(12, 16, arrmin[0])
        sheet1.write(12, 17, arravg[0])

        sheet1.write(13, 15, arrmax[1])
        sheet1.write(13, 16, arrmin[1])
        sheet1.write(13, 17, arravg[1])

        sheet1.write(14, 15, arrmax[2])
        sheet1.write(14, 16, arrmin[2])
        sheet1.write(14, 17, arravg[2])

        sheet1.write(15, 15, arrmax[3])
        sheet1.write(15, 16, arrmin[3])
        sheet1.write(15, 17, arravg[3])

        sheet1.write(16, 15, arrmax[4])
        sheet1.write(16, 16, arrmin[4])
        sheet1.write(16, 17, arravg[4])

        sheet1.write(17, 15, arrmax[5])
        sheet1.write(17, 16, arrmin[5])
        sheet1.write(17, 17, arravg[5])

        sheet1.write(18, 15, arrmax[6])
        sheet1.write(18, 16, arrmin[6])
        sheet1.write(18, 17, arravg[6])

        sheet1.write(19, 15, arrmax[7])
        sheet1.write(19, 16, arrmin[7])
        sheet1.write(19, 17, arravg[7])

        wb.save('HotSpot_Test_Results.xls')


if __name__ == "__main__":
    app = QApplication(sys.argv)

    w = MyForm()

    w.show()

    sys.exit(app.exec_())

