# import library
import tkinter as tk
import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output
import pandas as pd
import plotly.express as px


class TreeMapVisualization:

    def sealMain(self):

        HEIGHT = 500
        WIDTH = 600

        def btnClick():
            sealTreeMap()

        # GUI code
        # All code will will inside root & before mainloop()
        root = tk.Tk()

        root.title("Seal Treemap visualization")

        canvas = tk.Canvas(root, height=HEIGHT, width=WIDTH)
        canvas.pack()

        frame = tk.Frame(root, bg='#b3ffb3', bd=5)
        frame.place(relx=0.5, rely=0.02, relwidth=0.95, relheight=0.95, anchor='n')

        frame1 = tk.Frame(root, bg='white', bd=5)
        frame1.place(relx=0.5, rely=0.75, relwidth=0.65, relheight=0.1, anchor='n')

        lbl1 = tk.Label(frame, text='Treemap Graph', font=('comicsansms', 15, 'bold'))
        lbl1.place(relx=0.15, rely=0.05, relwidth=0.70, relheight=0.10)

        lbl2 = tk.Label(frame, text='Enter Excel \n file path: ', font=('comicsansms', 9, 'bold'))
        lbl2.place(relx=0.02, rely=0.2, relwidth=0.20, relheight=0.1)

        lower_frame = tk.Frame(root, bg='white', bd=5)
        lower_frame.place(relx=0.5, rely=0.75, relwidth=0.65, relheight=0.15, anchor='n')

        labelLowerFrame = tk.Label(lower_frame, bg='white', font=('comicsansms', 11, 'bold'))
        labelLowerFrame.place(relx=0.5, rely=0.1, relwidth=1, relheight=0.85, anchor='n')

        filePath = tk.StringVar(frame)
        filePathEntry = tk.Entry(frame, textvariable=filePath, font=('comicsansms', 8))
        filePathEntry.place(relx=0.25, rely=0.2, relwidth=0.72, relheight=0.1)

        btn = tk.Button(root, text="Generate Graph", bg='#ffa64d', command=btnClick, font=('comicsansms', 14, 'bold'))
        btn.place(relx=0.3, rely=0.40, relheight=0.1, relwidth=0.3)

        # Function for getting Excel file path
        def getExcelFilePath():
            excelFilePath = filePath.get()
            return excelFilePath

        # Plotly-Dash program starts from here
        def sealTreeMap():

            try:
                # Get the Excel Path
                excelPath = getExcelFilePath()

                # Read excel file
                df = pd.read_excel(excelPath, engine='openpyxl')

                sealApp = dash.Dash()

                # colors = {
                #     'background': '#111111',
                #     'text': '#000000'
                # }

                sealApp.layout = html.Div([
                    html.H4("Enter the JD Part Number or Supplier Part Number in below the input box",
                            style={"font-size": "14px", "font-family": "Roboto"}),
                    html.Span(["JD Part Number: ",
                               dcc.Input(id='my-input', value='', type='text')],
                              style={"margin-right": "40px", "font-size": "14px", "font-family": "Roboto"}),

                    html.Span(["Supplier Part Number: ",
                               dcc.Input(id='supplier-input', value='', type='text')],
                              style={"font-size": "14px", "font-family": "Roboto"}),

                    html.Div(
                        dcc.Graph(
                            id='graph-with-input'
                        )),

                ])

                @sealApp.callback(
                    Output(component_id='graph-with-input', component_property='figure'),
                    [Input(component_id='my-input', component_property='value'),
                     Input(component_id='supplier-input', component_property='value')]
                )
                def update_output_div(jdPartNumber, supplierPartNo):
                    if jdPartNumber == '' and supplierPartNo == '':
                        pass
                        return ''

                    else:
                        # Create copy of dataframe
                        dff = df.copy()

                        # Check whether John Deere Part Number and
                        # Supplier Part no. input box is empty or not
                        if jdPartNumber != '':
                            dff = dff[dff['partnumber'] == jdPartNumber]
                        elif supplierPartNo != '':
                            dff = dff[dff['supplierpartnumber'] == supplierPartNo]

                        # Store particular JD or Supplier part number parameter
                        jdPartNum = dff['partnumber']
                        groupName = dff['groupname']
                        groupName.dropna(inplace=True)
                        desription = dff['desription']
                        desription.dropna(inplace=True)
                        revision = dff['revision']
                        revision.dropna(inplace=True)
                        unitofmeasure = dff['unitofmeasure']
                        unitofmeasure.dropna(inplace=True)
                        includeonbom = dff['includeonbom']
                        includeonbom.dropna(inplace=True)
                        partStatus = dff['partstatus']
                        partStatus.dropna(inplace=True)
                        typecode = dff['typecode']
                        typecode.dropna(inplace=True)
                        colorcode = dff['colorcode']
                        colorcode.dropna(inplace=True)
                        materialcodeBase = dff['materialcode Base']
                        materialcodeBase.dropna(inplace=True)
                        architecturalcost = dff['architecturalcost']
                        architecturalcost.dropna(inplace=True)
                        replacedby = dff['replacedby']
                        replacedby.dropna(inplace=True)
                        alternatePartNumber = dff['alternatepartnumber']
                        alternatePartNumber.dropna(inplace=True)
                        weight = dff['weight']
                        weight.dropna(inplace=True)
                        specificationExtra = dff['specification extra']
                        specificationExtra.dropna(inplace=True)
                        userf1 = dff['userf1']
                        userf1.dropna(inplace=True)
                        userf2 = dff['userf2']
                        userf2.dropna(inplace=True)
                        partModified = dff['partmodified fomula']
                        partModified.dropna(inplace=True)
                        supplierPartNum = dff['supplierpartnumber']
                        supplierPartNum.dropna(inplace=True)
                        supplier = dff['suppliername']
                        supplier.dropna(inplace=True)
                        reelLength = dff['Reel Length']
                        reelLength.dropna(inplace=True)
                        supplierSpecification = dff['specification supplier']
                        supplierSpecification.dropna(inplace=True)
                        preferredSupplier = dff['preferred supplier']
                        preferredSupplier.dropna(inplace=True)
                        materialcodeST = dff['materialcode Single Terminations']
                        materialcodeST.dropna(inplace=True)
                        wireSpecST = dff['wirespec Single Terminations']
                        wireSpecST.dropna(inplace=True)
                        csaST = dff['csa Single Terminations']
                        csaST.dropna(inplace=True)
                        wireUsual = dff['wirusual']
                        wireUsual.dropna(inplace=True)
                        materialcodeMT = dff['materialcode Multiple Termination']
                        materialcodeMT.dropna(inplace=True)
                        wireSpecMT = dff['wirespec Multiple Termination']
                        wireSpecMT.dropna(inplace=True)
                        csaMT = dff['csa Multiple Termination']
                        csaMT.dropna(inplace=True)
                        property = dff['Userproperty Name']
                        property.dropna(inplace=True)
                        propDesc = dff['Prop Desc']
                        propDesc.dropna(inplace=True)
                        value = dff['Userproperty Value']
                        value.dropna(inplace=True)
                        msg = dff['msg']
                        msg.dropna(inplace=True)
                        modDate = dff['moddate fomula']
                        modDate.dropna(inplace=True)
                        modUser = dff['moduser']
                        modUser.dropna(inplace=True)

                        # replace emplty columns with ''
                        # str_cols = dff.columns[dff.dtypes == object]
                        # dff[str_cols] = df[str_cols].fillna('')
                        # dff.fillna('', inplace=True)

                        fig = px.treemap(dff,
                                         names=["John Deere Part Number: " + jdPartNum.iloc[0],
                                                "Supplier", "Supplier Part Number",
                                                ''.join([str(i) + '<br>' for i in supplierPartNum]),
                                                "Supplier Name", ''.join([str(i) + '<br>' for i in supplier]),
                                                "Reel Length",
                                                ''.join([str(i) + '<br>' for i in reelLength]),
                                                "Supplier Specification",
                                                ''.join([str(i) + '<br>' for i in supplierSpecification]),
                                                "Preferred", ''.join([str(i) + '<br>' for i in preferredSupplier]),
                                                "Base", "Group Name", groupName, "Desription", desription,
                                                "Revision", revision, "Unit Of Measure", unitofmeasure,
                                                "Include On BOM", includeonbom, "Status", partStatus, "Type Code",
                                                typecode,
                                                "Color Code", colorcode, "Material Code", materialcodeBase,
                                                "Extra", "Architectural Cost", architecturalcost, "Replaced By",
                                                replacedby, "Alternate Part", alternatePartNumber, "Weight", weight,
                                                "Specification", specificationExtra, "User Field1",
                                                userf1, "User Field2", userf2, "Last Modified",
                                                ''.join([str(i) + '<br>' for i in partModified]),
                                                "Properties", "Property", property, "Description ",
                                                propDesc, "Value", value,
                                                "Single Terminations", "Material Code ",
                                                ''.join([str(i + '<br>') for i in materialcodeST]),
                                                "Wire Specification", ''.join([str(i + '<br>') for i in wireSpecST]),
                                                "Wire CSA", ''.join([str(i) + '<br>' for i in csaST]), "Wire Usual",
                                                ''.join([str(i) + '<br>' for i in wireUsual]),
                                                "Multiple Terminations", "Material Code  ",
                                                ''.join([str(i + '<br>') for i in materialcodeMT]),
                                                "Wire Specification ", ''.join([str(i + '<br>') for i in wireSpecMT]),
                                                "History", "Msg", ''.join([str(i) + '<br>' for i in msg]),
                                                "ModDate", ''.join([str(i) + '<br>' for i in modDate]),
                                                "ModUser", ''.join([str(i) + '<br>' for i in modUser]),
                                                ],

                                         parents=["",
                                                  "John Deere Part Number: " + jdPartNum.iloc[0], "Supplier",
                                                  "Supplier Part Number",
                                                  "Supplier", "Supplier Name", "Supplier", "Reel Length",
                                                  "Supplier", "Supplier Specification", "Supplier", "Preferred",
                                                  "John Deere Part Number: " + jdPartNum.iloc[0], "Base", "Group Name",
                                                  "Base",
                                                  "Desription",
                                                  "Base", "Revision", "Base", "Unit Of Measure",
                                                  "Base", "Include On BOM", "Base", "Status", "Base", "Type Code",
                                                  "Base", "Color Code", "Base", "Material Code",
                                                  "John Deere Part Number: " + jdPartNum.iloc[0], "Extra",
                                                  "Architectural Cost",
                                                  "Extra",
                                                  "Replaced By", "Extra", "Alternate Part", "Extra", "Weight",
                                                  "Extra", "Specification", "Extra",
                                                  "User Field1", "Extra", "User Field2", "Extra", "Last Modified",
                                                  "John Deere Part Number: " + jdPartNum.iloc[0], "Properties",
                                                  "Property",
                                                  "Properties",
                                                  "Description ", "Properties", "Value",
                                                  "John Deere Part Number: " + jdPartNum.iloc[0], "Single Terminations",
                                                  "Material Code ",
                                                  "Single Terminations", "Wire Specification",
                                                  "Single Terminations", "Wire CSA", "Single Terminations",
                                                  "Wire Usual",
                                                  "John Deere Part Number: " + jdPartNum.iloc[0],
                                                  "Multiple Terminations",
                                                  "Material Code  ",
                                                  "Multiple Terminations", "Wire Specification ",
                                                  "John Deere Part Number: " + jdPartNum.iloc[0], "History", "Msg",
                                                  "History", "ModDate",
                                                  "History", "ModUser",
                                                  ],

                                         )

                        # Update/change layout
                        fig.update_layout(
                            title_font_size=200,
                            title_font_family='Roboto',
                            width=1400,
                            height=600,
                            margin=dict(l=20, r=10, t=30, b=10),
                            font=dict(
                                color="black",
                                size=12
                            ),
                        )
                        return fig

                # if __name__ == '__main__':
                sealApp.run_server(debug=False)

                labelLowerFrame['text'] = 'File converted'

            except Exception as e:
                labelLowerFrame['text'] = e
                print(e)

        root.mainloop()


if __name__ == '__main__':
    sealObj = TreeMapVisualization()
    sealObj.sealMain()

##################################################################################################


path = r'XML_To_Excel_Cavity_Seal_Nov_03.xlsm'
