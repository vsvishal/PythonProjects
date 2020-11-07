import os
import os.path
import tkinter as tk
import xml.etree.ElementTree as ET
from tkinter import ttk
import pandas as pd
import win32com.client   #pywin32


class ConnectorXMLToExcel():

    def connectorMain(self):

        HEIGHT = 500
        WIDTH = 600

        def btnClick():
            exportToExcel()

        # All code will will inside root & before mainloop()
        root = tk.Tk()

        root.title("Export Terminal or Connector data from XML")

        canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH)
        canvas.pack()

        frame = tk.Frame(root, bg='#b3ffb3', bd=5)
        frame.place(relx=0.5, rely=0.02, relwidth=0.95, relheight=0.95, anchor='n')

        frame1 = tk.Frame(root, bg='white', bd=5)
        frame1.place(relx=0.5, rely=0.75, relwidth=0.65, relheight=0.1, anchor='n')

        lbl1 = tk.Label(frame, text='Export XML to Excel Data', font=('comicsansms', 15, 'bold'))
        lbl1.place(relx=0.25, rely=0.01, relwidth=0.50, relheight=0.10)

        lbl2 = tk.Label(frame, text='XML file path: ', font=('comicsansms', 9, 'bold'))
        lbl2.place(relx=0.02, rely=0.15, relwidth=0.20, relheight=0.1)

        lbl3 = tk.Label(frame, text='Select Export type: ', font=('comicsansms', 9, 'bold'))
        lbl3.place(relx=0.02, rely=0.45, relwidth=0.20, relheight=0.1)

        lbl4 = tk.Label(frame, text='Excel file path: ', font=('comicsansms', 9, 'bold'))
        lbl4.place(relx=0.02, rely=0.30, relwidth=0.20, relheight=0.1)

        lower_frame = tk.Frame(root, bg='white', bd=5)
        lower_frame.place(relx=0.5, rely=0.75, relwidth=0.65, relheight=0.15, anchor='n')

        labelLowerFrame = tk.Label(lower_frame, bg='white', font=('comicsansms', 11, 'bold'))
        labelLowerFrame.place(relx=0.5, rely=0.1, relwidth=1, relheight=0.85, anchor='n')

        filePath = tk.StringVar(frame)
        filePathEntry = tk.Entry(frame, textvariable=filePath, font=('comicsansms', 8))
        filePathEntry.place(relx=0.25, rely=0.15, relwidth=0.72, relheight=0.1)

        excelfilePath = tk.StringVar(frame)
        excelfilePathEntry = tk.Entry(frame, textvariable=excelfilePath, font=('comicsansms', 8))
        excelfilePathEntry.place(relx=0.25, rely=0.30, relwidth=0.72, relheight=0.1)

        dropDwn = tk.StringVar()
        dropDwnChoosen = ttk.Combobox(frame, font=('comicsansms', 10), textvariable=dropDwn)
        # Adding combobox dropdown list
        dropDwnChoosen['values'] = ('Connector',
                                    'Terminal',
                                    'Seal')
        dropDwnChoosen.place(relx=0.25, rely=0.45, relwidth=0.72, relheight=0.1)
        dropDwnChoosen.current(1)

        btn = tk.Button(root, text="Convert XML to Excel", bg='#ffa64d', command=btnClick,
                        font=('comicsansms', 10, 'bold'))
        btn.place(relx=0.3, rely=0.60, relheight=0.1, relwidth=0.3)

        def getXmlPath():
            xmlPath = filePath.get()
            return xmlPath

        def getExcelPath():
           exportPath = excelfilePath.get()
           return exportPath

        def getDropDownVal():
            value = dropDwn.get()
            return value


        def exportToExcel():

            try:
                # Get the XML Path
                xmlPath = getXmlPath()
                # Get the Excel Path
                excelPath = getExcelPath()

                export_path_terminal = excelPath + r'\Terminal'
                export_path_connector = excelPath + r'\Connector'
                export_path_seal = excelPath + r'\Cavity_Seal'

                tree = ET.parse(xmlPath)
                root = tree.getroot()

                #################### Terminal #####################
                def terminal():

                    def terminalInnerDataExport(excel_file, sheet_name, df_cols):

                        data1 = []
                        data2 = []
                        data3 = []
                        data4 = []
                        data5 = []
                        data6 = []
                        data7 = []
                        data8 = []
                        data9 = []
                        data10 = []
                        data11 = []
                        data12 = []

                        for child in root:
                            for child_ele in child:
                                data1.append(child_ele.get(df_cols[0]))
                                data2.append(child_ele.get(df_cols[1]))
                                data3.append(child_ele.get(df_cols[2]))
                                data4.append(child_ele.get(df_cols[3]))
                                data5.append(child_ele.get(df_cols[4]))
                                data6.append(child_ele.get(df_cols[5]))
                                data7.append(child_ele.get(df_cols[6]))
                                data8.append(child_ele.get(df_cols[7]))
                                data9.append(child_ele.get(df_cols[8]))
                                data10.append(child_ele.get(df_cols[9]))
                                data11.append(child_ele.get(df_cols[10]))
                                data12.append(child_ele.get(df_cols[11]))

                        pd_data1 = pd.DataFrame(data1, columns=[df_cols[0]])
                        pd_data2 = pd.DataFrame(data2, columns=[df_cols[1]])
                        pd_data3 = pd.DataFrame(data3, columns=[df_cols[2]])
                        pd_data4 = pd.DataFrame(data4, columns=[df_cols[3]])
                        pd_data5 = pd.DataFrame(data5, columns=[df_cols[4]])
                        pd_data6 = pd.DataFrame(data6, columns=[df_cols[5]])
                        pd_data7 = pd.DataFrame(data7, columns=[df_cols[6]])
                        pd_data8 = pd.DataFrame(data8, columns=[df_cols[7]])
                        pd_data9 = pd.DataFrame(data9, columns=[df_cols[8]])
                        pd_data10 = pd.DataFrame(data10, columns=[df_cols[9]])
                        pd_data11 = pd.DataFrame(data11, columns=[df_cols[10]])
                        pd_data12 = pd.DataFrame(data12, columns=[df_cols[11]])

                        with pd.ExcelWriter(excel_file) as writer:
                            pd_data1.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0, index=False,
                                              na_rep='n.a')
                            pd_data2.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=1, index=False,
                                              na_rep='n.a')
                            pd_data3.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=2, index=False,
                                              na_rep='n.a')
                            pd_data4.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=3, index=False,
                                              na_rep='n.a')
                            pd_data5.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=4, index=False,
                                              na_rep='n.a')
                            pd_data6.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=5, index=False,
                                              na_rep='n.a')
                            pd_data7.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=6, index=False,
                                              na_rep='n.a')
                            pd_data8.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=7, index=False,
                                              na_rep='n.a')
                            pd_data9.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=8, index=False,
                                              na_rep='n.a')
                            pd_data10.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=9, index=False,
                                               na_rep='n.a')
                            pd_data11.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=10, index=False,
                                               na_rep='n.a')
                            pd_data12.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=11, index=False,
                                               na_rep='n.a')

                    def terminalDoubleInrDataExport(excel_file, sheet_name, df_cols):

                        data1 = []
                        data2 = []
                        data3 = []
                        data4 = []
                        data5 = []
                        data6 = []
                        data7 = []

                        for child in root:
                            for child_ele in child:
                                if child_ele.get(df_cols[3]) is not None:
                                    modificationHistoryId = child_ele.get(df_cols[3])
                                for child_ele_2 in child_ele:
                                    data1.append(child_ele_2.get(df_cols[0]))
                                    data2.append(child_ele_2.get(df_cols[1]))
                                    data3.append(child_ele_2.get(df_cols[2]))

                                    if child_ele_2.get(df_cols[4]) is not None:
                                        data4.append(modificationHistoryId)
                                        data5.append(child_ele_2.get(df_cols[4]))
                                        data6.append(child_ele_2.get(df_cols[5]))
                                        data7.append(child_ele_2.get(df_cols[6]))

                        pd_data1 = pd.DataFrame(data1, columns=[df_cols[0]])
                        pd_data2 = pd.DataFrame(data2, columns=[df_cols[1]])
                        pd_data3 = pd.DataFrame(data3, columns=[df_cols[2]])
                        pd_data4 = pd.DataFrame(data4, columns=[df_cols[3]])
                        pd_data5 = pd.DataFrame(data5, columns=[df_cols[4]])
                        pd_data6 = pd.DataFrame(data6, columns=[df_cols[5]])
                        pd_data7 = pd.DataFrame(data7, columns=[df_cols[6]])

                        with pd.ExcelWriter(excel_file) as writer:
                            pd_data1.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0, index=False,
                                              na_rep='n.a')
                            pd_data2.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=1, index=False,
                                              na_rep='n.a')
                            pd_data3.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=2, index=False,
                                              na_rep='n.a')
                            pd_data4.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=3, index=False,
                                              na_rep='n.a')
                            pd_data5.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=4, index=False,
                                              na_rep='n.a')
                            pd_data6.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=5, index=False,
                                              na_rep='n.a')
                            pd_data7.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=6, index=False,
                                              na_rep='n.a')

                    def terminalOuterDataExport(excel_file, sheet_name, df_cols):

                        data1 = []
                        data2 = []
                        data3 = []
                        data4 = []
                        data5 = []
                        data6 = []
                        data7 = []
                        data8 = []
                        data9 = []
                        data10 = []
                        data11 = []
                        data12 = []
                        data13 = []
                        data14 = []
                        data15 = []
                        data16 = []
                        data17 = []
                        data18 = []
                        data19 = []
                        data20 = []
                        data21 = []
                        data22 = []
                        data23 = []
                        data24 = []
                        data25 = []
                        data26 = []
                        data27 = []
                        data28 = []
                        data29 = []
                        data30 = []
                        data31 = []
                        data32 = []
                        data33 = []
                        data34 = []
                        data35 = []
                        data36 = []
                        data37 = []
                        data38 = []

                        for child in root:
                            data1.append(child.get(df_cols[0]))
                            data2.append(child.get(df_cols[1]))
                            data3.append(child.get(df_cols[2]))
                            data4.append(child.get(df_cols[3]))
                            data5.append(child.get(df_cols[4]))
                            data6.append(child.get(df_cols[5]))
                            data7.append(child.get(df_cols[6]))
                            data8.append(child.get(df_cols[7]))
                            data9.append(child.get(df_cols[8]))
                            data10.append(child.get(df_cols[9]))
                            data11.append(child.get(df_cols[10]))
                            data12.append(child.get(df_cols[11]))
                            data13.append(child.get(df_cols[12]))
                            data14.append(child.get(df_cols[13]))
                            data15.append(child.get(df_cols[14]))
                            data16.append(child.get(df_cols[15]))
                            data17.append(child.get(df_cols[16]))
                            data18.append(child.get(df_cols[17]))
                            data19.append(child.get(df_cols[18]))
                            data20.append(child.get(df_cols[19]))
                            data21.append(child.get(df_cols[20]))
                            data22.append(child.get(df_cols[21]))
                            data23.append(child.get(df_cols[22]))
                            data24.append(child.get(df_cols[23]))
                            data25.append(child.get(df_cols[24]))
                            data26.append(child.get(df_cols[25]))
                            data27.append(child.get(df_cols[26]))
                            data28.append(child.get(df_cols[27]))
                            data29.append(child.get(df_cols[28]))
                            data30.append(child.get(df_cols[29]))
                            data31.append(child.get(df_cols[30]))
                            data32.append(child.get(df_cols[31]))
                            data33.append(child.get(df_cols[32]))
                            data34.append(child.get(df_cols[33]))
                            data35.append(child.get(df_cols[34]))
                            data36.append(child.get(df_cols[35]))
                            data37.append(child.get(df_cols[36]))
                            data38.append(child.get(df_cols[37]))

                        pd_data1 = pd.DataFrame(data1, columns=[df_cols[0]])
                        pd_data2 = pd.DataFrame(data2, columns=[df_cols[1]])
                        pd_data3 = pd.DataFrame(data3, columns=[df_cols[2]])
                        pd_data4 = pd.DataFrame(data4, columns=[df_cols[3]])
                        pd_data5 = pd.DataFrame(data5, columns=[df_cols[4]])
                        pd_data6 = pd.DataFrame(data6, columns=[df_cols[5]])
                        pd_data7 = pd.DataFrame(data7, columns=[df_cols[6]])
                        pd_data8 = pd.DataFrame(data8, columns=[df_cols[7]])
                        pd_data9 = pd.DataFrame(data9, columns=[df_cols[8]])
                        pd_data10 = pd.DataFrame(data10, columns=[df_cols[9]])
                        pd_data11 = pd.DataFrame(data11, columns=[df_cols[10]])
                        pd_data12 = pd.DataFrame(data12, columns=[df_cols[11]])
                        pd_data13 = pd.DataFrame(data13, columns=[df_cols[12]])
                        pd_data14 = pd.DataFrame(data14, columns=[df_cols[13]])
                        pd_data15 = pd.DataFrame(data15, columns=[df_cols[14]])
                        pd_data16 = pd.DataFrame(data16, columns=[df_cols[15]])
                        pd_data17 = pd.DataFrame(data17, columns=[df_cols[16]])
                        pd_data18 = pd.DataFrame(data18, columns=[df_cols[17]])
                        pd_data19 = pd.DataFrame(data19, columns=[df_cols[18]])
                        pd_data20 = pd.DataFrame(data20, columns=[df_cols[19]])
                        pd_data21 = pd.DataFrame(data21, columns=[df_cols[20]])
                        pd_data22 = pd.DataFrame(data22, columns=[df_cols[21]])
                        pd_data23 = pd.DataFrame(data23, columns=[df_cols[22]])
                        pd_data24 = pd.DataFrame(data24, columns=[df_cols[23]])
                        pd_data25 = pd.DataFrame(data25, columns=[df_cols[24]])
                        pd_data26 = pd.DataFrame(data26, columns=[df_cols[25]])
                        pd_data27 = pd.DataFrame(data27, columns=[df_cols[26]])
                        pd_data28 = pd.DataFrame(data28, columns=[df_cols[27]])
                        pd_data29 = pd.DataFrame(data29, columns=[df_cols[28]])
                        pd_data30 = pd.DataFrame(data30, columns=[df_cols[29]])
                        pd_data31 = pd.DataFrame(data31, columns=[df_cols[30]])
                        pd_data32 = pd.DataFrame(data32, columns=[df_cols[31]])
                        pd_data33 = pd.DataFrame(data33, columns=[df_cols[32]])
                        pd_data34 = pd.DataFrame(data34, columns=[df_cols[33]])
                        pd_data35 = pd.DataFrame(data35, columns=[df_cols[34]])
                        pd_data36 = pd.DataFrame(data36, columns=[df_cols[35]])
                        pd_data37 = pd.DataFrame(data37, columns=[df_cols[36]])
                        pd_data38 = pd.DataFrame(data38, columns=[df_cols[37]])

                        with pd.ExcelWriter(excel_file) as writer:
                            pd_data1.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0, index=False,
                                              na_rep='n.a')
                            pd_data2.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=1, index=False,
                                              na_rep='n.a')
                            pd_data3.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=2, index=False,
                                              na_rep='n.a')
                            pd_data4.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=3, index=False,
                                              na_rep='n.a')
                            pd_data5.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=4, index=False,
                                              na_rep='n.a')
                            pd_data6.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=5, index=False,
                                              na_rep='n.a')
                            pd_data7.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=6, index=False,
                                              na_rep='n.a')
                            pd_data8.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=7, index=False,
                                              na_rep='n.a')
                            pd_data9.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=8, index=False,
                                              na_rep='n.a')
                            pd_data10.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=9, index=False,
                                               na_rep='n.a')
                            pd_data11.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=10, index=False,
                                               na_rep='n.a')
                            pd_data12.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=11, index=False,
                                               na_rep='n.a')
                            pd_data13.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=12, index=False,
                                               na_rep='n.a')
                            pd_data14.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=13, index=False,
                                               na_rep='n.a')
                            pd_data15.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=14, index=False,
                                               na_rep='n.a')
                            pd_data16.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=15, index=False,
                                               na_rep='n.a')
                            pd_data17.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=16, index=False,
                                               na_rep='n.a')
                            pd_data18.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=17, index=False,
                                               na_rep='n.a')
                            pd_data19.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=18, index=False,
                                               na_rep='n.a')
                            pd_data20.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=19, index=False,
                                               na_rep='n.a')
                            pd_data21.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=20, index=False,
                                               na_rep='n.a')
                            pd_data22.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=21, index=False,
                                               na_rep='n.a')
                            pd_data23.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=22, index=False,
                                               na_rep='n.a')
                            pd_data24.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=23, index=False,
                                               na_rep='n.a')
                            pd_data25.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=24, index=False,
                                               na_rep='n.a')
                            pd_data26.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=25, index=False,
                                               na_rep='n.a')
                            pd_data27.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=26, index=False,
                                               na_rep='n.a')
                            pd_data28.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=27, index=False,
                                               na_rep='n.a')
                            pd_data29.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=28, index=False,
                                               na_rep='n.a')
                            pd_data30.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=29, index=False,
                                               na_rep='n.a')
                            pd_data31.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=30, index=False,
                                               na_rep='n.a')
                            pd_data32.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=31, index=False,
                                               na_rep='n.a')
                            pd_data33.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=32, index=False,
                                               na_rep='n.a')
                            pd_data34.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=33, index=False,
                                               na_rep='n.a')
                            pd_data35.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=34, index=False,
                                               na_rep='n.a')
                            pd_data36.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=35, index=False,
                                               na_rep='n.a')
                            pd_data37.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=36, index=False,
                                               na_rep='n.a')
                            pd_data38.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=37, index=False,
                                               na_rep='n.a')

                    terminalInnerDataExport(export_path_terminal + '\Terminal_inner.xlsx', 'inner_PartData',
                                            ['libraryobject_id', 'supplierorganisation_id', 'supplierpartnumber_id',
                                             'librarysingletermination_id',
                                             'librarywirespec_id', 'librarymaterial_id',
                                             'librarymultipletermination_id', 'librarytermination_id',
                                             'supplierpartnumber', 'reellength', 'specification', 'preferred'])

                    terminalDoubleInrDataExport(export_path_terminal + '\Terminal_doubleInner.xlsx',
                                                'double_inside_data',
                                                ['librarymultipletermination_id', 'librarytermination_id',
                                                 'librarywirespec_id', 'libraryobject_id',
                                                 'msg', 'moddate', 'moduser'])

                    terminalOuterDataExport(export_path_terminal + '\Terminal_outer.xlsx', 'outside_data',
                                            ['libraryobject_id', 'librarywirespec_id', 'librarycomponenttype_id',
                                             'librarymaterial_id',
                                             'librarycolor_id', 'partnumber', 'description', 'revision', 'groupname',
                                             'materialcode', 'wirespec',
                                             'csa', 'wirusual', 'outsidediameter', 'typecode', 'colorcode',
                                             'unitofmeasure', 'architecturalcost',
                                             'cavityqt', 'includeonbom', 'partstatus', 'striplength', 'multstriplength',
                                             'addon', 'knockoff',
                                             'supplierorganisation_id', 'name', 'specification', 'replacedby',
                                             'alternatepartnumber', 'weight',
                                             'specification', 'userf1', 'userf2', 'userf3', 'userf4', 'userf5',
                                             'partmodified'])

                ################# Connector ###############
                def connector():

                    def connectorInnerData(excel_file, sheet_name, df_cols):

                        data1 = []
                        data2 = []
                        data3 = []
                        data4 = []
                        data5 = []
                        data6 = []
                        data7 = []
                        data8 = []
                        data9 = []
                        data10 = []
                        data11 = []
                        data12 = []
                        data13 = []
                        data14 = []
                        data15 = []
                        data16 = []
                        data17 = []
                        data18 = []
                        data19 = []
                        data20 = []
                        data21 = []
                        data22 = []
                        data23 = []
                        data24 = []
                        data25 = []
                        data26 = []
                        data27 = []
                        data28 = []
                        data29 = []
                        data30 = []
                        data31 = []
                        data32 = []
                        data33 = []
                        data34 = []
                        data35 = []
                        data36 = []
                        data37 = []
                        data38 = []
                        data39 = []
                        data40 = []
                        data41 = []
                        data42 = []
                        data43 = []
                        data44 = []
                        data45 = []
                        data46 = []
                        data47 = []
                        # data48 = []

                        for child in root:
                            for child_ele in child:
                                data1.append(child_ele.get(df_cols[0]))
                                data2.append(child_ele.get(df_cols[1]))
                                data3.append(child_ele.get(df_cols[2]))
                                data4.append(child_ele.get(df_cols[3]))
                                data5.append(child_ele.get(df_cols[4]))
                                data6.append(child_ele.get(df_cols[5]))
                                data7.append(child_ele.get(df_cols[6]))
                                data8.append(child_ele.get(df_cols[7]))
                                data9.append(child_ele.get(df_cols[8]))
                                data10.append(child_ele.get(df_cols[9]))
                                data11.append(child_ele.get(df_cols[10]))
                                data12.append(child_ele.get(df_cols[11]))
                                data13.append(child_ele.get(df_cols[12]))
                                data14.append(child_ele.get(df_cols[13]))
                                data15.append(child_ele.get(df_cols[14]))
                                data16.append(child_ele.get(df_cols[15]))
                                data17.append(child_ele.get(df_cols[16]))
                                data18.append(child_ele.get(df_cols[17]))
                                data19.append(child_ele.get(df_cols[18]))
                                data20.append(child_ele.get(df_cols[19]))
                                data21.append(child_ele.get(df_cols[20]))
                                data22.append(child_ele.get(df_cols[21]))
                                data23.append(child_ele.get(df_cols[22]))
                                data24.append(child_ele.get(df_cols[23]))
                                data25.append(child_ele.get(df_cols[24]))
                                data26.append(child_ele.get(df_cols[25]))
                                data27.append(child_ele.get(df_cols[26]))
                                data28.append(child_ele.get(df_cols[27]))
                                data29.append(child_ele.get(df_cols[28]))
                                data30.append(child_ele.get(df_cols[29]))
                                data31.append(child_ele.get(df_cols[30]))
                                data32.append(child_ele.get(df_cols[31]))
                                data33.append(child_ele.get(df_cols[32]))
                                data34.append(child_ele.get(df_cols[33]))
                                data35.append(child_ele.get(df_cols[34]))
                                data36.append(child_ele.get(df_cols[35]))
                                data37.append(child_ele.get(df_cols[36]))
                                data38.append(child_ele.get(df_cols[37]))
                                data39.append(child_ele.get(df_cols[38]))
                                data40.append(child_ele.get(df_cols[39]))
                                data41.append(child_ele.get(df_cols[40]))
                                data42.append(child_ele.get(df_cols[41]))
                                data43.append(child_ele.get(df_cols[42]))
                                data44.append(child_ele.get(df_cols[43]))
                                data45.append(child_ele.get(df_cols[44]))
                                data46.append(child_ele.get(df_cols[45]))
                                data47.append(child_ele.get(df_cols[46]))
                                # data48.append(child_ele.get(df_cols[47]))

                        pd_data1 = pd.DataFrame(data1, columns=[df_cols[0]])
                        pd_data2 = pd.DataFrame(data2, columns=[df_cols[1]])
                        pd_data3 = pd.DataFrame(data3, columns=[df_cols[2]])
                        pd_data4 = pd.DataFrame(data4, columns=[df_cols[3]])
                        pd_data5 = pd.DataFrame(data5, columns=[df_cols[4]])
                        pd_data6 = pd.DataFrame(data6, columns=[df_cols[5]])
                        pd_data7 = pd.DataFrame(data7, columns=[df_cols[6]])
                        pd_data8 = pd.DataFrame(data8, columns=[df_cols[7]])
                        pd_data9 = pd.DataFrame(data9, columns=[df_cols[8]])
                        pd_data10 = pd.DataFrame(data10, columns=[df_cols[9]])
                        pd_data11 = pd.DataFrame(data11, columns=[df_cols[10]])
                        pd_data12 = pd.DataFrame(data12, columns=[df_cols[11]])
                        pd_data13 = pd.DataFrame(data13, columns=[df_cols[12]])
                        pd_data14 = pd.DataFrame(data14, columns=[df_cols[13]])
                        pd_data15 = pd.DataFrame(data15, columns=[df_cols[14]])
                        pd_data16 = pd.DataFrame(data16, columns=[df_cols[15]])
                        pd_data17 = pd.DataFrame(data17, columns=[df_cols[16]])
                        pd_data18 = pd.DataFrame(data18, columns=[df_cols[17]])
                        pd_data19 = pd.DataFrame(data19, columns=[df_cols[18]])
                        pd_data20 = pd.DataFrame(data20, columns=[df_cols[19]])
                        pd_data21 = pd.DataFrame(data21, columns=[df_cols[20]])
                        pd_data22 = pd.DataFrame(data22, columns=[df_cols[21]])
                        pd_data23 = pd.DataFrame(data23, columns=[df_cols[22]])
                        pd_data24 = pd.DataFrame(data24, columns=[df_cols[23]])
                        pd_data25 = pd.DataFrame(data25, columns=[df_cols[24]])
                        pd_data26 = pd.DataFrame(data26, columns=[df_cols[25]])
                        pd_data27 = pd.DataFrame(data27, columns=[df_cols[26]])
                        pd_data28 = pd.DataFrame(data28, columns=[df_cols[27]])
                        pd_data29 = pd.DataFrame(data29, columns=[df_cols[28]])
                        pd_data30 = pd.DataFrame(data30, columns=[df_cols[29]])
                        pd_data31 = pd.DataFrame(data31, columns=[df_cols[30]])
                        pd_data32 = pd.DataFrame(data32, columns=[df_cols[31]])
                        pd_data33 = pd.DataFrame(data33, columns=[df_cols[32]])
                        pd_data34 = pd.DataFrame(data34, columns=[df_cols[33]])
                        pd_data35 = pd.DataFrame(data35, columns=[df_cols[34]])
                        pd_data36 = pd.DataFrame(data36, columns=[df_cols[35]])
                        pd_data37 = pd.DataFrame(data37, columns=[df_cols[36]])
                        pd_data38 = pd.DataFrame(data38, columns=[df_cols[37]])
                        pd_data39 = pd.DataFrame(data39, columns=[df_cols[38]])
                        pd_data40 = pd.DataFrame(data40, columns=[df_cols[39]])
                        pd_data41 = pd.DataFrame(data41, columns=[df_cols[40]])
                        pd_data42 = pd.DataFrame(data42, columns=[df_cols[41]])
                        pd_data43 = pd.DataFrame(data43, columns=[df_cols[42]])
                        pd_data44 = pd.DataFrame(data44, columns=[df_cols[43]])
                        pd_data45 = pd.DataFrame(data45, columns=[df_cols[44]])
                        pd_data46 = pd.DataFrame(data46, columns=[df_cols[45]])
                        pd_data47 = pd.DataFrame(data47, columns=[df_cols[46]])
                        # pd_data48 = pd.DataFrame(data48, columns=[df_cols[47]])

                        with pd.ExcelWriter(excel_file) as writer:
                            pd_data1.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0, index=False,
                                              na_rep='n.a')
                            pd_data2.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=1, index=False,
                                              na_rep='n.a')
                            pd_data3.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=2, index=False,
                                              na_rep='n.a')
                            pd_data4.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=3, index=False,
                                              na_rep='n.a')
                            pd_data5.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=4, index=False,
                                              na_rep='n.a')
                            pd_data6.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=5, index=False,
                                              na_rep='n.a')
                            pd_data7.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=6, index=False,
                                              na_rep='n.a')
                            pd_data8.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=7, index=False,
                                              na_rep='n.a')
                            pd_data9.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=8, index=False,
                                              na_rep='n.a')
                            pd_data10.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=9, index=False,
                                               na_rep='n.a')
                            pd_data11.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=10, index=False,
                                               na_rep='n.a')
                            pd_data12.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=11, index=False,
                                               na_rep='n.a')
                            pd_data13.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=12, index=False,
                                               na_rep='n.a')
                            pd_data14.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=13, index=False,
                                               na_rep='n.a')
                            pd_data15.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=14, index=False,
                                               na_rep='n.a')
                            pd_data16.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=15, index=False,
                                               na_rep='n.a')
                            pd_data17.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=16, index=False,
                                               na_rep='n.a')
                            pd_data18.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=17, index=False,
                                               na_rep='n.a')
                            pd_data19.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=18, index=False,
                                               na_rep='n.a')
                            pd_data20.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=19, index=False,
                                               na_rep='n.a')
                            pd_data21.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=20, index=False,
                                               na_rep='n.a')
                            pd_data22.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=21, index=False,
                                               na_rep='n.a')
                            pd_data23.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=22, index=False,
                                               na_rep='n.a')
                            pd_data24.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=23, index=False,
                                               na_rep='n.a')
                            pd_data25.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=24, index=False,
                                               na_rep='n.a')
                            pd_data26.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=25, index=False,
                                               na_rep='n.a')
                            pd_data27.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=26, index=False,
                                               na_rep='n.a')
                            pd_data28.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=27, index=False,
                                               na_rep='n.a')
                            pd_data29.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=28, index=False,
                                               na_rep='n.a')
                            pd_data30.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=29, index=False,
                                               na_rep='n.a')
                            pd_data31.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=30, index=False,
                                               na_rep='n.a')
                            pd_data32.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=31, index=False,
                                               na_rep='n.a')
                            pd_data33.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=32, index=False,
                                               na_rep='n.a')
                            pd_data34.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=33, index=False,
                                               na_rep='n.a')
                            pd_data35.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=34, index=False,
                                               na_rep='n.a')
                            pd_data36.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=35, index=False,
                                               na_rep='n.a')
                            pd_data37.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=36, index=False,
                                               na_rep='n.a')
                            pd_data38.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=37, index=False,
                                               na_rep='n.a')
                            pd_data39.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=38, index=False,
                                               na_rep='n.a')
                            pd_data40.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=39, index=False,
                                               na_rep='n.a')
                            pd_data41.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=40, index=False,
                                               na_rep='n.a')
                            pd_data42.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=41, index=False,
                                               na_rep='n.a')
                            pd_data43.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=42, index=False,
                                               na_rep='n.a')
                            pd_data44.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=43, index=False,
                                               na_rep='n.a')
                            pd_data45.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=44, index=False,
                                               na_rep='n.a')
                            pd_data46.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=45, index=False,
                                               na_rep='n.a')
                            pd_data47.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=46, index=False,
                                               na_rep='n.a')
                            # pd_data48.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=47, index=False, na_rep='n.a')

                        ########################################### Connector doubleInner  ##########################################################

                    def connectorDoubleInner(excel_file, sheet_name, df_cols):

                        data1 = []
                        data2 = []
                        data3 = []
                        data4 = []
                        data5 = []
                        data6 = []
                        data7 = []
                        data8 = []
                        data9 = []
                        data10 = []

                        for child in root:

                            for child_ele in child:
                                for child_ele_2 in child_ele:
                                    data1.append(child_ele_2.get(df_cols[0]))
                                    data2.append(child_ele_2.get(df_cols[1]))
                                    data3.append(child_ele_2.get(df_cols[2]))
                                    data4.append(child_ele_2.get(df_cols[3]))
                                    data5.append(child_ele_2.get(df_cols[4]))
                                    data6.append(child_ele_2.get(df_cols[5]))
                                    data7.append(child_ele_2.get(df_cols[6]))
                                    data8.append(child_ele_2.get(df_cols[7]))
                                    data9.append(child_ele_2.get(df_cols[8]))
                                    data10.append(child_ele_2.get(df_cols[9]))

                        pd_data1 = pd.DataFrame(data1, columns=[df_cols[0]])
                        pd_data2 = pd.DataFrame(data2, columns=[df_cols[1]])
                        pd_data3 = pd.DataFrame(data3, columns=[df_cols[2]])
                        pd_data4 = pd.DataFrame(data4, columns=[df_cols[3]])
                        pd_data5 = pd.DataFrame(data5, columns=[df_cols[4]])
                        pd_data6 = pd.DataFrame(data6, columns=[df_cols[5]])
                        pd_data7 = pd.DataFrame(data7, columns=[df_cols[6]])
                        pd_data8 = pd.DataFrame(data8, columns=[df_cols[7]])
                        pd_data9 = pd.DataFrame(data9, columns=[df_cols[8]])
                        pd_data10 = pd.DataFrame(data10, columns=[df_cols[9]])

                        with pd.ExcelWriter(excel_file) as writer:
                            pd_data1.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0, index=False,
                                              na_rep='n.a')
                            pd_data2.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=1, index=False,
                                              na_rep='n.a')
                            pd_data3.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=2, index=False,
                                              na_rep='n.a')
                            pd_data4.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=3, index=False,
                                              na_rep='n.a')
                            pd_data5.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=4, index=False,
                                              na_rep='n.a')
                            pd_data6.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=5, index=False,
                                              na_rep='n.a')
                            pd_data7.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=6, index=False,
                                              na_rep='n.a')
                            pd_data8.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=7, index=False,
                                              na_rep='n.a')
                            pd_data9.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=8, index=False,
                                              na_rep='n.a')
                            pd_data10.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=9, index=False,
                                               na_rep='n.a')

                    ################################# outside_data  ############################################
                    def connectorOuterData(excel_file, sheet_name, df_cols):

                        data1 = []
                        data2 = []
                        data3 = []
                        data4 = []
                        data5 = []
                        data6 = []
                        data7 = []
                        data8 = []
                        data9 = []
                        data10 = []
                        data11 = []
                        data12 = []
                        data13 = []
                        data14 = []
                        data15 = []
                        data16 = []
                        data17 = []
                        data18 = []
                        data19 = []
                        data20 = []
                        data21 = []
                        data22 = []
                        data23 = []
                        data24 = []
                        data25 = []
                        data26 = []
                        data27 = []
                        data28 = []
                        data29 = []
                        data30 = []
                        data31 = []
                        data32 = []
                        data33 = []
                        data34 = []
                        data35 = []
                        data36 = []
                        data37 = []
                        data38 = []
                        data39 = []
                        data40 = []
                        data41 = []
                        data42 = []
                        data43 = []

                        for child in root:
                            data1.append(child.get(df_cols[0]))
                            data2.append(child.get(df_cols[1]))
                            data3.append(child.get(df_cols[2]))
                            data4.append(child.get(df_cols[3]))
                            data5.append(child.get(df_cols[4]))
                            data6.append(child.get(df_cols[5]))
                            data7.append(child.get(df_cols[6]))
                            data8.append(child.get(df_cols[7]))
                            data9.append(child.get(df_cols[8]))
                            data10.append(child.get(df_cols[9]))
                            data11.append(child.get(df_cols[10]))
                            data12.append(child.get(df_cols[11]))
                            data13.append(child.get(df_cols[12]))
                            data14.append(child.get(df_cols[13]))
                            data15.append(child.get(df_cols[14]))
                            data16.append(child.get(df_cols[15]))
                            data17.append(child.get(df_cols[16]))
                            data18.append(child.get(df_cols[17]))
                            data19.append(child.get(df_cols[18]))
                            data20.append(child.get(df_cols[19]))
                            data21.append(child.get(df_cols[20]))
                            data22.append(child.get(df_cols[21]))
                            data23.append(child.get(df_cols[22]))
                            data24.append(child.get(df_cols[23]))
                            data25.append(child.get(df_cols[24]))
                            data26.append(child.get(df_cols[25]))
                            data27.append(child.get(df_cols[26]))
                            data28.append(child.get(df_cols[27]))
                            data29.append(child.get(df_cols[28]))
                            data30.append(child.get(df_cols[29]))
                            data31.append(child.get(df_cols[30]))
                            data32.append(child.get(df_cols[31]))
                            data33.append(child.get(df_cols[32]))
                            data34.append(child.get(df_cols[33]))
                            data35.append(child.get(df_cols[34]))
                            data36.append(child.get(df_cols[35]))
                            data37.append(child.get(df_cols[36]))
                            data38.append(child.get(df_cols[37]))
                            data39.append(child.get(df_cols[38]))
                            data40.append(child.get(df_cols[39]))
                            data41.append(child.get(df_cols[40]))
                            data42.append(child.get(df_cols[41]))
                            data43.append(child.get(df_cols[42]))

                        pd_data1 = pd.DataFrame(data1, columns=[df_cols[0]])
                        pd_data2 = pd.DataFrame(data2, columns=[df_cols[1]])
                        pd_data3 = pd.DataFrame(data3, columns=[df_cols[2]])
                        pd_data4 = pd.DataFrame(data4, columns=[df_cols[3]])
                        pd_data5 = pd.DataFrame(data5, columns=[df_cols[4]])
                        pd_data6 = pd.DataFrame(data6, columns=[df_cols[5]])
                        pd_data7 = pd.DataFrame(data7, columns=[df_cols[6]])
                        pd_data8 = pd.DataFrame(data8, columns=[df_cols[7]])
                        pd_data9 = pd.DataFrame(data9, columns=[df_cols[8]])
                        pd_data10 = pd.DataFrame(data10, columns=[df_cols[9]])
                        pd_data11 = pd.DataFrame(data11, columns=[df_cols[10]])
                        pd_data12 = pd.DataFrame(data12, columns=[df_cols[11]])
                        pd_data13 = pd.DataFrame(data13, columns=[df_cols[12]])
                        pd_data14 = pd.DataFrame(data14, columns=[df_cols[13]])
                        pd_data15 = pd.DataFrame(data15, columns=[df_cols[14]])
                        pd_data16 = pd.DataFrame(data16, columns=[df_cols[15]])
                        pd_data17 = pd.DataFrame(data17, columns=[df_cols[16]])
                        pd_data18 = pd.DataFrame(data18, columns=[df_cols[17]])
                        pd_data19 = pd.DataFrame(data19, columns=[df_cols[18]])
                        pd_data20 = pd.DataFrame(data20, columns=[df_cols[19]])
                        pd_data21 = pd.DataFrame(data21, columns=[df_cols[20]])
                        pd_data22 = pd.DataFrame(data22, columns=[df_cols[21]])
                        pd_data23 = pd.DataFrame(data23, columns=[df_cols[22]])
                        pd_data24 = pd.DataFrame(data24, columns=[df_cols[23]])
                        pd_data25 = pd.DataFrame(data25, columns=[df_cols[24]])
                        pd_data26 = pd.DataFrame(data26, columns=[df_cols[25]])
                        pd_data27 = pd.DataFrame(data27, columns=[df_cols[26]])
                        pd_data28 = pd.DataFrame(data28, columns=[df_cols[27]])
                        pd_data29 = pd.DataFrame(data29, columns=[df_cols[28]])
                        pd_data30 = pd.DataFrame(data30, columns=[df_cols[29]])
                        pd_data31 = pd.DataFrame(data31, columns=[df_cols[30]])
                        pd_data32 = pd.DataFrame(data32, columns=[df_cols[31]])
                        pd_data33 = pd.DataFrame(data33, columns=[df_cols[32]])
                        pd_data34 = pd.DataFrame(data34, columns=[df_cols[33]])
                        pd_data35 = pd.DataFrame(data35, columns=[df_cols[34]])
                        pd_data36 = pd.DataFrame(data36, columns=[df_cols[35]])
                        pd_data37 = pd.DataFrame(data37, columns=[df_cols[36]])
                        pd_data38 = pd.DataFrame(data38, columns=[df_cols[37]])
                        pd_data39 = pd.DataFrame(data39, columns=[df_cols[38]])
                        pd_data40 = pd.DataFrame(data40, columns=[df_cols[39]])
                        pd_data41 = pd.DataFrame(data41, columns=[df_cols[40]])
                        pd_data42 = pd.DataFrame(data42, columns=[df_cols[41]])
                        pd_data43 = pd.DataFrame(data43, columns=[df_cols[42]])

                        with pd.ExcelWriter(excel_file) as writer:
                            pd_data1.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0, index=False,
                                              na_rep='n.a')
                            pd_data2.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=1, index=False,
                                              na_rep='n.a')
                            pd_data3.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=2, index=False,
                                              na_rep='n.a')
                            pd_data4.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=3, index=False,
                                              na_rep='n.a')
                            pd_data5.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=4, index=False,
                                              na_rep='n.a')
                            pd_data6.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=5, index=False,
                                              na_rep='n.a')
                            pd_data7.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=6, index=False,
                                              na_rep='n.a')
                            pd_data8.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=7, index=False,
                                              na_rep='n.a')
                            pd_data9.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=8, index=False,
                                              na_rep='n.a')
                            pd_data10.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=9, index=False,
                                               na_rep='n.a')
                            pd_data11.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=10, index=False,
                                               na_rep='n.a')
                            pd_data12.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=11, index=False,
                                               na_rep='n.a')
                            pd_data13.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=12, index=False,
                                               na_rep='n.a')
                            pd_data14.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=13, index=False,
                                               na_rep='n.a')
                            pd_data15.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=14, index=False,
                                               na_rep='n.a')
                            pd_data16.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=15, index=False,
                                               na_rep='n.a')
                            pd_data17.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=16, index=False,
                                               na_rep='n.a')
                            pd_data18.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=17, index=False,
                                               na_rep='n.a')
                            pd_data19.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=18, index=False,
                                               na_rep='n.a')
                            pd_data20.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=19, index=False,
                                               na_rep='n.a')
                            pd_data21.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=20, index=False,
                                               na_rep='n.a')
                            pd_data22.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=21, index=False,
                                               na_rep='n.a')
                            pd_data23.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=22, index=False,
                                               na_rep='n.a')
                            pd_data24.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=23, index=False,
                                               na_rep='n.a')
                            pd_data25.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=24, index=False,
                                               na_rep='n.a')
                            pd_data26.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=25, index=False,
                                               na_rep='n.a')
                            pd_data27.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=26, index=False,
                                               na_rep='n.a')
                            pd_data28.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=27, index=False,
                                               na_rep='n.a')
                            pd_data29.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=28, index=False,
                                               na_rep='n.a')
                            pd_data30.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=29, index=False,
                                               na_rep='n.a')
                            pd_data31.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=30, index=False,
                                               na_rep='n.a')
                            pd_data32.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=31, index=False,
                                               na_rep='n.a')
                            pd_data33.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=32, index=False,
                                               na_rep='n.a')
                            pd_data34.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=33, index=False,
                                               na_rep='n.a')
                            pd_data35.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=34, index=False,
                                               na_rep='n.a')
                            pd_data36.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=35, index=False,
                                               na_rep='n.a')
                            pd_data37.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=36, index=False,
                                               na_rep='n.a')
                            pd_data38.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=37, index=False,
                                               na_rep='n.a')
                            pd_data39.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=38, index=False,
                                               na_rep='n.a')
                            pd_data40.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=39, index=False,
                                               na_rep='n.a')
                            pd_data41.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=40, index=False,
                                               na_rep='n.a')
                            pd_data42.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=41, index=False,
                                               na_rep='n.a')
                            pd_data43.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=42, index=False,
                                               na_rep='n.a')

                    connectorInnerData(
                        export_path_connector + '\Connector_inner.xlsx',
                        'inner_PartData',
                        ['libraryobject_id', 'supplierorganisation_id', 'supplierpartnumber_id',
                         'librarysingletermination_id',
                         'librarysinglewire_id', 'librarywirespec_id', 'librarymaterial_id', 'symbol_id',
                         'housdef_id', 'housingdefinition_id', 'librarycavity_id', 'librarypincontainer_id',
                         'cavitygroup_id', 'libraryrevision_id', 'revisiongrp_id',
                         'chsuserpropertypart_id', 'chsuserproperty_id', 'librarydressedroute_id',
                         'librarymating_id', 'matedconnector_id',
                         'librarydevicefootprint_id', 'mappedcavity_id', 'subcomponen_id',
                         'supplierpartnumber', 'reellength', 'specification',
                         'preferred', 'quantity', 'position', 'selectionstatus', 'psblocked', 'mode', 'scope',
                         'priority',
                         'userpropertyvalue', 'userpropertyname', 'propdesc', 'cavityname', 'sortorder', 'isblocked',
                         'cavity', 'librarygraphic_id', 'context', 'graphiccode', 'route', 'wire_addon',
                         'wire_knockoff'])

                    connectorDoubleInner(
                        export_path_connector + '\Connector_doubleInner.xlsx',
                        'double_inside_data',
                        ['libraryobject_id', 'librarymultipletermination_id', 'librarytermination_id',
                         'librarywirespec_id',
                         'property_id', 'librarymatingpinmapping_id', 'librarymating_id', 'librarycavity_id',
                         'librarydevicefootprint_id', 'mappedcavity_id',
                         ])

                    connectorOuterData(
                        export_path_connector + '\Connector_outer.xlsx',
                        'outside_data',
                        ['libraryobject_id', 'librarywirespec_id', 'librarycomponenttype_id', 'librarymaterial_id',
                         'librarycolor_id', 'librarycomponenttype_id', 'revisiongrp_id',
                         'chsuserproperty_id', 'partnumber', 'description', 'revision', 'groupname', 'materialcode',
                         'wirespec',
                         'csa', 'wirusual', 'outsidediameter', 'typecode', 'colorcode', 'unitofmeasure',
                         'architecturalcost',
                         'cavityqt', 'includeonbom', 'partstatus', 'striplength', 'multstriplength', 'addon',
                         'knockoff',
                         'supplierorganisation_id', 'name', 'specification', 'replacedby', 'alternatepartnumber',
                         'weight',
                         'specification', 'userf1', 'userf2', 'userf3', 'userf4', 'userf5', 'partmodified',
                         'userpropertyname', 'propdesc'])

                #############################Seal ###########################
                def seal():

                    ##################### Seal Inner Data ##############
                    def sealInnerData(excel_file, sheet_name, df_cols):

                        data1 = []
                        data2 = []
                        data3 = []
                        data4 = []
                        data5 = []
                        data6 = []
                        data7 = []
                        data8 = []
                        data9 = []
                        data10 = []
                        data11 = []
                        data12 = []
                        data13 = []
                        data14 = []
                        data15 = []
                        data16 = []

                        for child in root:
                            for child_ele in child:
                                data1.append(child_ele.get(df_cols[0]))
                                data2.append(child_ele.get(df_cols[1]))
                                data3.append(child_ele.get(df_cols[2]))
                                data4.append(child_ele.get(df_cols[3]))
                                data5.append(child_ele.get(df_cols[4]))
                                data6.append(child_ele.get(df_cols[5]))
                                data7.append(child_ele.get(df_cols[6]))
                                data8.append(child_ele.get(df_cols[7]))
                                data9.append(child_ele.get(df_cols[8]))
                                data10.append(child_ele.get(df_cols[9]))
                                data11.append(child_ele.get(df_cols[10]))
                                data12.append(child_ele.get(df_cols[11]))
                                data13.append(child_ele.get(df_cols[12]))
                                data14.append(child_ele.get(df_cols[13]))
                                data15.append(child_ele.get(df_cols[14]))
                                data16.append(child_ele.get(df_cols[15]))

                        pd_data1 = pd.DataFrame(data1, columns=[df_cols[0]])
                        pd_data2 = pd.DataFrame(data2, columns=[df_cols[1]])
                        pd_data3 = pd.DataFrame(data3, columns=[df_cols[2]])
                        pd_data4 = pd.DataFrame(data4, columns=[df_cols[3]])
                        pd_data5 = pd.DataFrame(data5, columns=[df_cols[4]])
                        pd_data6 = pd.DataFrame(data6, columns=[df_cols[5]])
                        pd_data7 = pd.DataFrame(data7, columns=[df_cols[6]])
                        pd_data8 = pd.DataFrame(data8, columns=[df_cols[7]])
                        pd_data9 = pd.DataFrame(data9, columns=[df_cols[8]])
                        pd_data10 = pd.DataFrame(data10, columns=[df_cols[9]])
                        pd_data11 = pd.DataFrame(data11, columns=[df_cols[10]])
                        pd_data12 = pd.DataFrame(data12, columns=[df_cols[11]])
                        pd_data13 = pd.DataFrame(data13, columns=[df_cols[12]])
                        pd_data14 = pd.DataFrame(data14, columns=[df_cols[13]])
                        pd_data15 = pd.DataFrame(data15, columns=[df_cols[14]])
                        pd_data16 = pd.DataFrame(data16, columns=[df_cols[15]])

                        with pd.ExcelWriter(excel_file) as writer:
                            pd_data1.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0, index=False,
                                              na_rep='n.a')
                            pd_data2.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=1, index=False,
                                              na_rep='n.a')
                            pd_data3.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=2, index=False,
                                              na_rep='n.a')
                            pd_data4.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=3, index=False,
                                              na_rep='n.a')
                            pd_data5.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=4, index=False,
                                              na_rep='n.a')
                            pd_data6.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=5, index=False,
                                              na_rep='n.a')
                            pd_data7.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=6, index=False,
                                              na_rep='n.a')
                            pd_data8.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=7, index=False,
                                              na_rep='n.a')
                            pd_data9.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=8, index=False,
                                              na_rep='n.a')
                            pd_data10.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=9, index=False,
                                               na_rep='n.a')
                            pd_data11.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=10, index=False,
                                               na_rep='n.a')
                            pd_data12.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=11, index=False,
                                               na_rep='n.a')
                            pd_data13.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=12, index=False,
                                               na_rep='n.a')
                            pd_data14.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=13, index=False,
                                               na_rep='n.a')
                            pd_data15.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=14, index=False,
                                               na_rep='n.a')
                            pd_data16.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=15, index=False,
                                               na_rep='n.a')

                    ################## Seal Double_Inner Data ################3
                    def sealDoubleInner(excel_file, sheet_name, df_cols):

                        data1 = []
                        data2 = []
                        data3 = []
                        data4 = []
                        data5 = []
                        data6 = []
                        data7 = []

                        for child in root:
                            for child_ele in child:
                                if child_ele.get(df_cols[3]) is not None:
                                    modificationHistoryId = child_ele.get(df_cols[3])
                                for child_ele_2 in child_ele:
                                    data1.append(child_ele_2.get(df_cols[0]))
                                    data2.append(child_ele_2.get(df_cols[1]))
                                    data3.append(child_ele_2.get(df_cols[2]))

                                    if child_ele_2.get(df_cols[4]) is not None:
                                        data4.append(modificationHistoryId)
                                        data5.append(child_ele_2.get(df_cols[4]))
                                        data6.append(child_ele_2.get(df_cols[5]))
                                        data7.append(child_ele_2.get(df_cols[6]))

                        pd_data1 = pd.DataFrame(data1, columns=[df_cols[0]])
                        pd_data2 = pd.DataFrame(data2, columns=[df_cols[1]])
                        pd_data3 = pd.DataFrame(data3, columns=[df_cols[2]])
                        pd_data4 = pd.DataFrame(data4, columns=[df_cols[3]])
                        pd_data5 = pd.DataFrame(data5, columns=[df_cols[4]])
                        pd_data6 = pd.DataFrame(data6, columns=[df_cols[5]])
                        pd_data7 = pd.DataFrame(data7, columns=[df_cols[6]])

                        with pd.ExcelWriter(excel_file) as writer:
                            pd_data1.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0, index=False,
                                              na_rep='n.a')
                            pd_data2.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=1, index=False,
                                              na_rep='n.a')
                            pd_data3.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=2, index=False,
                                              na_rep='n.a')
                            pd_data4.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=3, index=False,
                                              na_rep='n.a')
                            pd_data5.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=4, index=False,
                                              na_rep='n.a')
                            pd_data6.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=5, index=False,
                                              na_rep='n.a')
                            pd_data7.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=6, index=False,
                                              na_rep='n.a')

                    ###################### Seal Outer Data ###################
                    def sealOuterData(excel_file, sheet_name, df_cols):

                        data1 = []
                        data2 = []
                        data3 = []
                        data4 = []
                        data5 = []
                        data6 = []
                        data7 = []
                        data8 = []
                        data9 = []
                        data10 = []
                        data11 = []
                        data12 = []
                        data13 = []
                        data14 = []
                        data15 = []
                        data16 = []
                        data17 = []
                        data18 = []
                        data19 = []
                        data20 = []
                        data21 = []
                        data22 = []
                        data23 = []
                        data24 = []
                        data25 = []
                        data26 = []
                        data27 = []
                        data28 = []
                        data29 = []
                        data30 = []
                        data31 = []
                        data32 = []
                        data33 = []
                        data34 = []
                        data35 = []
                        data36 = []
                        data37 = []
                        data38 = []
                        data39 = []
                        data40 = []
                        data41 = []
                        data42 = []

                        for child in root:
                            data1.append(child.get(df_cols[0]))
                            data2.append(child.get(df_cols[1]))
                            data3.append(child.get(df_cols[2]))
                            data4.append(child.get(df_cols[3]))
                            data5.append(child.get(df_cols[4]))
                            data6.append(child.get(df_cols[5]))
                            data7.append(child.get(df_cols[6]))
                            data8.append(child.get(df_cols[7]))
                            data9.append(child.get(df_cols[8]))
                            data10.append(child.get(df_cols[9]))
                            data11.append(child.get(df_cols[10]))
                            data12.append(child.get(df_cols[11]))
                            data13.append(child.get(df_cols[12]))
                            data14.append(child.get(df_cols[13]))
                            data15.append(child.get(df_cols[14]))
                            data16.append(child.get(df_cols[15]))
                            data17.append(child.get(df_cols[16]))
                            data18.append(child.get(df_cols[17]))
                            data19.append(child.get(df_cols[18]))
                            data20.append(child.get(df_cols[19]))
                            data21.append(child.get(df_cols[20]))
                            data22.append(child.get(df_cols[21]))
                            data23.append(child.get(df_cols[22]))
                            data24.append(child.get(df_cols[23]))
                            data25.append(child.get(df_cols[24]))
                            data26.append(child.get(df_cols[25]))
                            data27.append(child.get(df_cols[26]))
                            data28.append(child.get(df_cols[27]))
                            data29.append(child.get(df_cols[28]))
                            data30.append(child.get(df_cols[29]))
                            data31.append(child.get(df_cols[30]))
                            data32.append(child.get(df_cols[31]))
                            data33.append(child.get(df_cols[32]))
                            data34.append(child.get(df_cols[33]))
                            data35.append(child.get(df_cols[34]))
                            data36.append(child.get(df_cols[35]))
                            data37.append(child.get(df_cols[36]))
                            data38.append(child.get(df_cols[37]))
                            data39.append(child.get(df_cols[38]))
                            data40.append(child.get(df_cols[39]))
                            data41.append(child.get(df_cols[40]))
                            data42.append(child.get(df_cols[41]))

                        pd_data1 = pd.DataFrame(data1, columns=[df_cols[0]])
                        pd_data2 = pd.DataFrame(data2, columns=[df_cols[1]])
                        pd_data3 = pd.DataFrame(data3, columns=[df_cols[2]])
                        pd_data4 = pd.DataFrame(data4, columns=[df_cols[3]])
                        pd_data5 = pd.DataFrame(data5, columns=[df_cols[4]])
                        pd_data6 = pd.DataFrame(data6, columns=[df_cols[5]])
                        pd_data7 = pd.DataFrame(data7, columns=[df_cols[6]])
                        pd_data8 = pd.DataFrame(data8, columns=[df_cols[7]])
                        pd_data9 = pd.DataFrame(data9, columns=[df_cols[8]])
                        pd_data10 = pd.DataFrame(data10, columns=[df_cols[9]])
                        pd_data11 = pd.DataFrame(data11, columns=[df_cols[10]])
                        pd_data12 = pd.DataFrame(data12, columns=[df_cols[11]])
                        pd_data13 = pd.DataFrame(data13, columns=[df_cols[12]])
                        pd_data14 = pd.DataFrame(data14, columns=[df_cols[13]])
                        pd_data15 = pd.DataFrame(data15, columns=[df_cols[14]])
                        pd_data16 = pd.DataFrame(data16, columns=[df_cols[15]])
                        pd_data17 = pd.DataFrame(data17, columns=[df_cols[16]])
                        pd_data18 = pd.DataFrame(data18, columns=[df_cols[17]])
                        pd_data19 = pd.DataFrame(data19, columns=[df_cols[18]])
                        pd_data20 = pd.DataFrame(data20, columns=[df_cols[19]])
                        pd_data21 = pd.DataFrame(data21, columns=[df_cols[20]])
                        pd_data22 = pd.DataFrame(data22, columns=[df_cols[21]])
                        pd_data23 = pd.DataFrame(data23, columns=[df_cols[22]])
                        pd_data24 = pd.DataFrame(data24, columns=[df_cols[23]])
                        pd_data25 = pd.DataFrame(data25, columns=[df_cols[24]])
                        pd_data26 = pd.DataFrame(data26, columns=[df_cols[25]])
                        pd_data27 = pd.DataFrame(data27, columns=[df_cols[26]])
                        pd_data28 = pd.DataFrame(data28, columns=[df_cols[27]])
                        pd_data29 = pd.DataFrame(data29, columns=[df_cols[28]])
                        pd_data30 = pd.DataFrame(data30, columns=[df_cols[29]])
                        pd_data31 = pd.DataFrame(data31, columns=[df_cols[30]])
                        pd_data32 = pd.DataFrame(data32, columns=[df_cols[31]])
                        pd_data33 = pd.DataFrame(data33, columns=[df_cols[32]])
                        pd_data34 = pd.DataFrame(data34, columns=[df_cols[33]])
                        pd_data35 = pd.DataFrame(data35, columns=[df_cols[34]])
                        pd_data36 = pd.DataFrame(data36, columns=[df_cols[35]])
                        pd_data37 = pd.DataFrame(data37, columns=[df_cols[36]])
                        pd_data38 = pd.DataFrame(data38, columns=[df_cols[37]])
                        pd_data39 = pd.DataFrame(data39, columns=[df_cols[38]])
                        pd_data40 = pd.DataFrame(data40, columns=[df_cols[39]])
                        pd_data41 = pd.DataFrame(data41, columns=[df_cols[40]])
                        pd_data42 = pd.DataFrame(data42, columns=[df_cols[41]])

                        with pd.ExcelWriter(excel_file) as writer:
                            pd_data1.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=0, index=False,
                                              na_rep='n.a')
                            pd_data2.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=1, index=False,
                                              na_rep='n.a')
                            pd_data3.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=2, index=False,
                                              na_rep='n.a')
                            pd_data4.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=3, index=False,
                                              na_rep='n.a')
                            pd_data5.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=4, index=False,
                                              na_rep='n.a')
                            pd_data6.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=5, index=False,
                                              na_rep='n.a')
                            pd_data7.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=6, index=False,
                                              na_rep='n.a')
                            pd_data8.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=7, index=False,
                                              na_rep='n.a')
                            pd_data9.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=8, index=False,
                                              na_rep='n.a')
                            pd_data10.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=9, index=False,
                                               na_rep='n.a')
                            pd_data11.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=10, index=False,
                                               na_rep='n.a')
                            pd_data12.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=11, index=False,
                                               na_rep='n.a')
                            pd_data13.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=12, index=False,
                                               na_rep='n.a')
                            pd_data14.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=13, index=False,
                                               na_rep='n.a')
                            pd_data15.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=14, index=False,
                                               na_rep='n.a')
                            pd_data16.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=15, index=False,
                                               na_rep='n.a')
                            pd_data17.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=16, index=False,
                                               na_rep='n.a')
                            pd_data18.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=17, index=False,
                                               na_rep='n.a')
                            pd_data19.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=18, index=False,
                                               na_rep='n.a')
                            pd_data20.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=19, index=False,
                                               na_rep='n.a')
                            pd_data21.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=20, index=False,
                                               na_rep='n.a')
                            pd_data22.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=21, index=False,
                                               na_rep='n.a')
                            pd_data23.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=22, index=False,
                                               na_rep='n.a')
                            pd_data24.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=23, index=False,
                                               na_rep='n.a')
                            pd_data25.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=24, index=False,
                                               na_rep='n.a')
                            pd_data26.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=25, index=False,
                                               na_rep='n.a')
                            pd_data27.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=26, index=False,
                                               na_rep='n.a')
                            pd_data28.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=27, index=False,
                                               na_rep='n.a')
                            pd_data29.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=28, index=False,
                                               na_rep='n.a')
                            pd_data30.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=29, index=False,
                                               na_rep='n.a')
                            pd_data31.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=30, index=False,
                                               na_rep='n.a')
                            pd_data32.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=31, index=False,
                                               na_rep='n.a')
                            pd_data33.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=32, index=False,
                                               na_rep='n.a')
                            pd_data34.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=33, index=False,
                                               na_rep='n.a')
                            pd_data35.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=34, index=False,
                                               na_rep='n.a')
                            pd_data36.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=35, index=False,
                                               na_rep='n.a')
                            pd_data37.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=36, index=False,
                                               na_rep='n.a')
                            pd_data38.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=37, index=False,
                                               na_rep='n.a')
                            pd_data39.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=38, index=False,
                                               na_rep='n.a')
                            pd_data40.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=39, index=False,
                                               na_rep='n.a')
                            pd_data41.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=40, index=False,
                                               na_rep='n.a')
                            pd_data42.to_excel(writer, sheet_name=sheet_name, startrow=0, startcol=41, index=False,
                                               na_rep='n.a')

                    sealInnerData(export_path_seal + '\Seal_inner.xlsx', 'inner_PartData',
                                  ['libraryobject_id', 'supplierorganisation_id', 'supplierpartnumber_id',
                                   'librarysingletermination_id',
                                   'librarywirespec_id', 'librarymaterial_id', 'librarymultipletermination_id',
                                   'librarytermination_id', 'chsuserproperty_id', 'userpropertyvalue',
                                   'userpropertyname', 'propdesc',
                                   'supplierpartnumber', 'reellength', 'specification', 'preferred'])

                    sealDoubleInner(export_path_seal + '\Seal_doubleInner.xlsx',
                                    'double_inside_data',
                                    ['librarymultipletermination_id', 'librarytermination_id', 'librarywirespec_id',
                                     'libraryobject_id', 'msg', 'moddate', 'moduser'
                                     ])

                    sealOuterData(export_path_seal + '\Seal_outer.xlsx', 'outside_data',
                                  ['libraryobject_id', 'librarywirespec_id', 'librarycomponenttype_id',
                                   'librarymaterial_id',
                                   'librarycolor_id', 'chsuserproperty_id', 'partnumber', 'description', 'revision',
                                   'groupname',
                                   'materialcode', 'wirespec',
                                   'csa', 'wirusual', 'outsidediameter', 'typecode', 'colorcode', 'unitofmeasure',
                                   'architecturalcost',
                                   'cavityqt', 'includeonbom', 'partstatus', 'striplength', 'multstriplength', 'addon',
                                   'knockoff',
                                   'supplierorganisation_id', 'name', 'specification', 'replacedby',
                                   'alternatepartnumber', 'weight',
                                   'specification', 'userf1', 'userf2', 'userf3', 'userf4', 'userf5', 'partmodified',
                                   'userpropertyvalue', 'userpropertyname', 'propdesc', ])

                # Check whether drop down selection was Terminal or Connector
                dropDownValueCheck = getDropDownVal()
                if dropDownValueCheck == "Terminal":
                    terminal()
                    # Call Connector vba macro sub proceude
                    if os.path.exists(
                            export_path_terminal + "\XML_To_Excel_Terminal_all_wires.xlsm"):
                        xl = win32com.client.Dispatch("Excel.Application")
                        xl.Workbooks.Open(os.path.abspath(
                            export_path_terminal + "\XML_To_Excel_Terminal_all_wires.xlsm"),
                            ReadOnly=1)
                        xl.Application.Run("XML_To_Excel_Terminal_all_wires.xlsm!Main.TerminalMain")
                        # xl.Application.Save() # if you want to save then uncomment this line and change delete the
                        # ", ReadOnly=1" part from the open function.
                        xl.Application.Quit()  # Comment this out if your excel script closes
                        del xl

                elif dropDownValueCheck == "Connector":
                    connector()
                    # Call Connector vba macro sub proceude
                    if os.path.exists(
                            export_path_connector + "\XML_To_Excel_Connector.xlsm"):
                        xl = win32com.client.Dispatch("Excel.Application")
                        xl.Workbooks.Open(os.path.abspath(
                            export_path_connector + "\XML_To_Excel_Connector.xlsm"),
                            ReadOnly=1)
                        xl.Application.Run("XML_To_Excel_Connector.xlsm!Main.ConnectorMain")
                        # xl.Application.Save() # if you want to save then uncomment this line and change delete the
                        # ", ReadOnly=1" part from the open function.
                        xl.Application.Quit()  # Comment this out if your excel script closes
                        del xl

                elif dropDownValueCheck == "Seal":
                    seal()
                    # Call Seal vba macro sub proceude
                    if os.path.exists(
                            export_path_seal + "\XML_To_Excel_Cavity_Seal.xlsm"):
                        xl = win32com.client.Dispatch("Excel.Application")
                        xl.Workbooks.Open(os.path.abspath(
                            export_path_seal + "\XML_To_Excel_Cavity_Seal.xlsm"),
                            ReadOnly=1)
                        xl.Application.Run("XML_To_Excel_Cavity_Seal.xlsm!Main.SealMain")
                        # xl.Application.Save() # if you want to save then uncomment this line and change delete the
                        # ", ReadOnly=1" part from the open function.
                        xl.Application.Quit()  # Comment this out if your excel script closes
                        del xl

                else:
                    labelLowerFrame['text'] = "Wrong drop down option"

                labelLowerFrame['text'] = 'File converted'

            except Exception as e:
                labelLowerFrame['text'] = e
                print(e)

        root.mainloop()


if __name__ == '__main__':
    xmlObj = ConnectorXMLToExcel()
    xmlObj.connectorMain()
#######################################
