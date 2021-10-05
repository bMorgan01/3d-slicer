import os
import sys
from copy import copy, deepcopy
from math import sqrt, ceil, floor
import numpy as np
import tkinter as tk
from tkinter import filedialog, LEFT, X
from rectpack import newPacker, float2dec, PackingMode, PackingBin, SORT_NONE
from ctypes import windll

GWL_EXSTYLE = -20
WS_EX_APPWINDOW = 0x00040000
WS_EX_TOOLWINDOW = 0x00000080

import trimesh
from trimesh.path.entities import Text
from trimesh.path.exchange.export import export_path
from trimesh.path.creation import rectangle
from trimesh.transformations import scale_matrix

# declare "globals" for use across functions
mesh = None
unit = None
sections = None
pdfSections = None
combined = None
bins = None
blocked = False
saveBinsButton = None
showKeyButton = None
exportHeightEntry = None
exportWidthEntry = None
exportKerfEntry = None
lastW = None
lastH = None
lastK = None
resultWinHeight = None
resultWinWidth = None


# shows dimensions and layer count in main gui window
# updates live when model has been selected
def calcDim(args=None):
    scale = 1

    # if scale is number set scale, else reset text entry
    try:
        scale = float(scaleFactorEntry.get())
    except ValueError:
        scaleFactorEntry.config(text="1")

    # get and display scaled bounds
    dimensions = mesh.bounds

    dimensionsAfterLabel.config(text="Scaled Dimensions (" + unit + ")")
    xAfterlabel.config(text="x: " + str(round((dimensions[1][0] - dimensions[0][0]) * scale, 2)))
    yAfterlabel.config(text="y: " + str(round((dimensions[1][1] - dimensions[0][1]) * scale, 2)))
    zAfterlabel.config(text="z: " + str(round((dimensions[1][2] - dimensions[0][2]) * scale, 2)))

    # get and display number of layers
    try:
        thickness = float(layerThicknessEntry.get())
        if thickness == 0:
            numLayersLabel.config(text="Layers: N/A")
        else:
            numLayersLabel.config(
                text="Layers: " + str(ceil(((dimensions[1][2] - dimensions[0][2]) * scale) / thickness)))
    except ValueError:
        numLayersLabel.config(text="Layers: N/A")


# opens 3D model and fills gui elements with appropriate info (base unit, model name, model dimensions, etc)
def getFile():
    # maps base unit result to abbreviation
    unitDict = {
        "millimeters": "mm",
        "inches": "in"
    }

    # ask for file
    filename = tk.filedialog.askopenfilename(initialdir="/",
                                             title="Select a File",
                                             filetypes=((".stl files", "*.stl"), ("all files", ".*")))

    # add model file path to gui
    modelEntry.config(text=filename)

    # load the mesh from filename
    global mesh, unit
    mesh = trimesh.load_mesh(filename)

    # get base unit, add to gui
    unit = trimesh.units.units_from_metadata(mesh, guess=True)
    try:
        unit = unitDict[unit]
    except KeyError:
        pass

    layerThicknessLabel.config(text="Layer Thickness (" + unit + "): ")

    # add model dimensions to gui
    calcDim()

    # now that model is open, listen for changes in relevant text entries to update dimensions shown in gui
    scaleFactorEntry.bind('<KeyRelease>', calcDim)
    layerThicknessEntry.bind('<KeyRelease>', calcDim)


def focus_results():
    # bring results window forward
    if results is not None and 'normal' == results.state():
        results.focus_set()
        return


# run slices, show export window
def go():
    global combined, sections, results, saveBinsButton, showKeyButton, showKey, exportHeightEntry, exportWidthEntry, exportKerfEntry, resultWinWidth, resultWinHeight

    focus_results()

    # scale model according to text entry in main window
    mesh.apply_transform(scale_matrix(float(scaleFactorEntry.get()), [0, 0, 0]))

    # slice the mesh into evenly spaced chunks along z
    # this takes the (2,3) bounding box and slices it into [minz, maxz]
    z_extents = mesh.bounds[:, 2]
    # slice every x (text entry) model units (eg, inches)
    z_levels = np.arange(*z_extents, step=float(layerThicknessEntry.get()))

    # create cross section outlines
    sections = mesh.section_multiplane(plane_origin=mesh.bounds[0],
                                       plane_normal=[0, 0, 1],
                                       heights=z_levels)

    # in the case that slice did not intersect with model, remove section
    sections = list(filter(lambda element: element is not None, sections))

    combined = np.sum(sections)

    # GUI stuff
    results = tk.Toplevel(root)
    results.title("Results")
    results.protocol("WM_DELETE_WINDOW", clearResults)
    results.resizable(0, 0)

    buttonFrame = tk.Frame(results)
    viewFrame = tk.Frame(buttonFrame)
    sectionsButton = tk.Button(viewFrame, text="View Layers", command=plot)
    modelButton = tk.Button(viewFrame, text="View Model", command=model)
    exportButton = tk.Button(buttonFrame, text="Export Layers to SVG", command=export)

    heightFrame = tk.Frame(results)
    exportHeight = tk.Label(heightFrame, text="Height (" + unit + "): ")
    exportHeightEntry = tk.Entry(heightFrame, width=10)

    widthFrame = tk.Frame(results)
    exportWidth = tk.Label(widthFrame, text="Width (" + unit + "): ")
    exportWidthEntry = tk.Entry(widthFrame, width=10)

    kerfFrame = tk.Frame(results)
    exportKerf = tk.Label(kerfFrame, text="Kerf (" + unit + "): ")
    exportKerfEntry = tk.Entry(kerfFrame, width=10)

    exportFileButton = tk.Button(results, text="Prepare Cuts", command=(
        lambda *args: exportFile(float(exportHeightEntry.get()), float(exportWidthEntry.get()),
                                 float(exportKerfEntry.get()))))

    exportFrame = tk.Frame(results)
    saveBinsButton = tk.Button(exportFrame, text="Export Cut Files", command=saveFiles)
    showKeyButton = tk.Checkbutton(exportFrame, text='Key?', variable=showKey, onvalue=True, offvalue=False)

    buttonFrame.pack(pady=10, padx=5)
    viewFrame.pack(pady=5)
    sectionsButton.pack(side=LEFT, padx=1)
    modelButton.pack(side=LEFT, padx=1)
    exportButton.pack()

    heightFrame.pack()
    exportHeight.pack(side=LEFT)
    exportHeightEntry.pack(side=LEFT)
    widthFrame.pack()
    exportWidth.pack(side=LEFT)
    exportWidthEntry.pack(side=LEFT)
    kerfFrame.pack()
    exportKerf.pack(side=LEFT)
    exportKerfEntry.pack(side=LEFT)

    exportFileButton.pack(padx=1)
    exportFrame.pack(pady=5, padx=5)

    results.update()
    results.geometry(f'+{root.winfo_rootx()+int((root.winfo_width()-results.winfo_width())/2)}+{root.winfo_rooty()+int((root.winfo_height()-results.winfo_height())/2)}')

    resultWinWidth = results.winfo_width()
    resultWinHeight = results.winfo_height()
    results.mainloop()


# empty results, hide window so process can be run again
def clearResults():
    global results

    results.destroy()
    results = None


# show top-down view of stacked layers
def plot():
    global blocked

    if not blocked:
        blocked = True
        # summing the array of Path2D objects will put all of the curves
        # into one Path2D object, which can be plotted easily
        combined.show()

        blocked = False


# shows unsliced model
def model():
    global blocked

    if not blocked:
        blocked = True

        scene = mesh.scene()
        camera = scene.camera
        camera.resolution = (640, 520)
        scene.base_frame = "Model"
        scene.camera = camera
        scene.show()

        blocked = False


# save layer svg's to selected directory
def export():
    global blocked

    if not blocked:
        blocked = True

        # ask for directory
        directory = filedialog.askdirectory()

        for i in range(len(sections)):
            export_path(sections[i], file_type="svg", file_obj=directory + "/layer" + str(i) + ".svg")

        focus_results()

        blocked = False


def checkPrepared(args=None):
    if float(exportWidthEntry.get()) != lastW or float(exportHeightEntry.get()) != lastH or float(exportKerfEntry.get()) != lastK:
        saveBinsButton.pack_forget()
        showKeyButton.pack_forget()

        results.geometry(f"{resultWinHeight}x{resultWinHeight}")
    else:
        if not saveBinsButton.winfo_ismapped():
            results.geometry("")
            saveBinsButton.pack(side=LEFT, padx=1)
            showKeyButton.pack(side=LEFT)

# prepares bins for export, performs bin packing
# arguments: bin height, bin width, kerf
def exportFile(h, w, k):
    global blocked, bins, pdfSections, exportWidthEntry, exportHeightEntry, exportKerfEntry, lastW, lastH, lastK

    # if other window is not open
    if not blocked:
        blocked = True

        lastW = w
        lastH = h
        lastK = k

        # prepare packers, one with rotation enabled, one disabled, results will be compared later
        packer = newPacker(mode=PackingMode.Offline, bin_algo=PackingBin.Global, sort_algo=SORT_NONE, rotation=False)
        rotPacker = newPacker(mode=PackingMode.Offline, bin_algo=PackingBin.Global, sort_algo=SORT_NONE, rotation=True)

        # add rectangles to packers
        for i in range(len(sections)):
            dimensions = sections[i].bounds

            sect = deepcopy(sections[i])
            sect.apply_translation((-dimensions[0][0], -dimensions[0][1]))

            dimensions = sect.bounds

            # set dimensions to bounds plus kerf
            r = (float2dec((dimensions[1][0]) + k, 5), float2dec((dimensions[1][1]) + k, 5))
            packer.add_rect(*r, i)
            rotPacker.add_rect(*r, i)

        # add bins to packers
        packer.add_bin(float2dec(w, 5), float2dec(h, 5), count=float("inf"))
        rotPacker.add_bin(float2dec(w, 5), float2dec(h, 5), count=float("inf"))

        # run pack
        packer.pack()
        rotPacker.pack()

        # init var to largest bin area of packers, will be overridden to find minimum area used, minimizing material wastage
        leastBinSum = w * h * max(len(packer), len(rotPacker))
        leastPacker = packer
        # loop through packers
        for pack in (packer, rotPacker):
            binSum = 0
            # loop through packed bins
            for i in range(len(pack)):
                leastX = w
                leastY = h
                greatestX = 0
                greatestY = 0
                # loop through packed rectangles
                for rect in pack.rect_list():
                    # if rectangles is in bin
                    if rect[0] == i:
                        # check to see if rectangle is closer to the top corner than the previous closest
                        if rect[1] < leastX:
                            leastX = rect[1]
                        if rect[2] < leastY:
                            leastY = rect[2]
                        # check to see if rectangle is closer to the bottom corner than the previous closest
                        if rect[3] + rect[1] > greatestX:
                            greatestX = rect[3] + rect[1]
                        if rect[4] + rect[2] > greatestY:
                            greatestY = rect[4] + rect[2]
                # add used area of bin to total
                binSum += ((greatestX - leastX) * (greatestY - leastY))
            # override packer least used area if needed
            if binSum < leastBinSum:
                leastBinSum = binSum
                leastPacker = pack

        # use algo that resulted in the least wastage
        packer = leastPacker
        bins = list()
        displaySections = list()
        pdfSections = list()
        # create visual representation of bins for display
        for i in range(len(packer)):
            bins.append(list())

            rowColLen = ceil(sqrt(len(packer)))
            row = floor(i / rowColLen)
            col = i % rowColLen

            displayBin = rectangle(((0, 0), (w, h)))
            displayBin.apply_translation(
                ((packer[i].width * 11 / 10) * col, -((packer[i].height * 11 / 10) * row + packer[i].height)))

            displaySections.append(displayBin)
            pdfSections.append(displayBin)

        # add rectangles to their designated bins
        for rect in packer.rect_list():
            section = deepcopy(sections[rect[5]])
            bin = rect[0]

            # find position of bin for display purposes (bins displayed in grid, find row & col)
            rowColLen = ceil(sqrt(len(packer)))
            row = floor((bin) / rowColLen)
            col = bin % rowColLen

            # check if rectangle has rotated, if so, rotate
            dimensions = section.bounds
            if rect[3] != float2dec((dimensions[1][0]) + k, 5) and rect[4] != float2dec((dimensions[1][1]) + k, 5):
                section.apply_transform(((0, -1, 0), (1, 0, 0), (0, 0, 1)))
                dimensions = section.bounds
                section.apply_translation((-dimensions[0][0], -dimensions[0][1]))

            # find x, y translation to move rectangle to designated spot in bin
            section.apply_translation((rect[1] + float2dec(k / 2, 5), rect[2] + float2dec(k / 2, 5)))
            bins[bin].append(section)

            # find x, y translation to offset display rectangle to the proper bin in the bin grid
            displaySection = copy(section)
            displaySection.apply_translation(
                ((packer[bin].width * 11 / 10) * col, -((packer[bin].height * 11 / 10) * row + packer[bin].height)))
            displaySections.append(displaySection)

            # for display only, add rectangle bounding box, for visualization of packing performance
            displayRect = rectangle(((displaySection.bounds[0][0], displaySection.bounds[0][1]),
                                     (displaySection.bounds[1][0], displaySection.bounds[1][1])))
            displaySections.append(displayRect)

            # also add bounding box to key display, real shape outline not added for readability
            pdfRect = copy(displayRect)
            pdfRect.entities = np.append(pdfRect.entities, Text(0, str(rect[5] + 1), align=('left', 'bottom')))
            pdfSections.append(pdfRect)

        # show packed rectangles and bins in human understandable format
        displayCombine = np.sum(displaySections)
        displayCombine.show()

        # now that export has been prepared, show save buttons
        saveBinsButton.pack(side=LEFT, padx=1)
        showKeyButton.pack(side=LEFT)

        exportWidthEntry.bind('<KeyRelease>', checkPrepared)
        exportHeightEntry.bind('<KeyRelease>', checkPrepared)
        exportKerfEntry.bind('<KeyRelease>', checkPrepared)

        blocked = False


# save bins and show key
def saveFiles():
    global blocked

    if not blocked:
        blocked = True

        # ask for directory selection
        directory = filedialog.askdirectory()

        # export bins
        for i in range(len(bins)):
            exportCombine = np.sum(bins[i])
            export_path(exportCombine, file_type="svg", file_obj=directory + "/bin" + str(i) + ".svg")

        focus_results()

        # show key
        if (showKey.get()):
            pdfCombine = np.sum(pdfSections)
            pdfCombine.show()

        blocked = False


# forces window to appear in taskbar, windows specific
def set_appwindow(root):
    hwnd = windll.user32.GetParent(root.winfo_id())
    style = windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
    style = style & ~WS_EX_TOOLWINDOW
    style = style | WS_EX_APPWINDOW
    res = windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
    # re-assert the new window style
    root.wm_withdraw()
    root.after(10, lambda: root.wm_deiconify())


# for newer versions of pyinstaller, needed to retrieve icon image
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


results = None

# create base window
root = tk.Tk()
root.title("3DSlicer")
root.option_add('*Font', '19')
root.attributes("-toolwindow", 1)
p1 = tk.PhotoImage(file=resource_path('icon.png'))
# Icon set for program window
root.iconphoto(True, p1)
root.resizable(0, 0)

# clear children of base window
clearList = root.winfo_children()

for child in clearList:
    child.destroy()

# set default state of key checkbox
showKey = tk.BooleanVar()

# generate window
menubar = tk.Menu(root)
filemenu = tk.Menu(menubar, tearoff=0)

# TODO: filemenu.add_command(label="About", command=about)
filemenu.add_command(label="Exit", command=root.destroy)
menubar.add_cascade(label="File", menu=filemenu)

modelFrame = tk.Frame(root)
topModelFrame = tk.Frame(modelFrame)
modelLabel = tk.Label(topModelFrame, text="Model: ")
modelEntry = tk.Label(topModelFrame, text="N/A")
modelBrowse = tk.Button(modelFrame, text="Browse", command=getFile)

entriesFrame = tk.Frame(root)
layerEntryFrame = tk.Frame(entriesFrame)
layerThicknessLabel = tk.Label(layerEntryFrame, text="Layer Thickness (N/A): ")
layerThicknessEntry = tk.Entry(layerEntryFrame, width=5)

scaleEntryFrame = tk.Frame(entriesFrame)
scaleFactorLabel = tk.Label(scaleEntryFrame, text="Model Scale Factor: ")
scaleFactorEntry = tk.Entry(scaleEntryFrame, width=5)
scaleFactorEntry.insert(0, "1")

numbersFrame = tk.Frame(root)
dimensionsAfterLabel = tk.Label(numbersFrame, text="Scaled Dimensions (N/A)")
xAfterlabel = tk.Label(numbersFrame, text="x: N/A")
yAfterlabel = tk.Label(numbersFrame, text="y: N/A")
zAfterlabel = tk.Label(numbersFrame, text="z: N/A")

numLayersLabel = tk.Label(root, text="Layers: N/A")

goButton = tk.Button(root, text="Go", command=go)

root.config(menu=menubar)

modelFrame.pack(pady=10)
topModelFrame.pack()
modelLabel.pack(side=LEFT)
modelEntry.pack(side=LEFT)
modelBrowse.pack()

entriesFrame.pack()
numbersFrame.pack(pady=10)
layerEntryFrame.pack()
scaleEntryFrame.pack(fill=X)
scaleFactorLabel.pack(side=LEFT)
scaleFactorEntry.pack(side=LEFT)
layerThicknessLabel.pack(side=LEFT)
layerThicknessEntry.pack(side=LEFT)

dimensionsAfterLabel.pack()
xAfterlabel.pack()
yAfterlabel.pack()
zAfterlabel.pack()

numLayersLabel.pack()

goButton.pack(ipadx=10, pady=5)

root.after(10, lambda: set_appwindow(root))
root.mainloop()
