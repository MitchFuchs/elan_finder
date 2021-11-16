"""
Created on June 2021
@author: MitchFuchs
@contact: michael@neptune-consulting.ch

Parse .eaf files (XML) and allow to access video clips directly from ELAN annotations
"""

from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import cv2
import PIL.Image, PIL.ImageTk
import os, sys, subprocess
import xml.etree.ElementTree as ET
import json
import pandas as pd
import xlsxwriter as xw
from datetime import datetime
import time


class ElanFinder:

    def __init__(self):
        self.root = Tk()
        self.root.configure(background='light grey')
        self.root.title("ELAN finder")

        self.annotations = {}
        self.time_slots = {}
        self.reversed_time_slots = {}

        self.all_tiers = []
        self.annotation_values = []
        self.relevant_annotation_values = []
        self.meta_tiers = ["Context", "Phases", "Subphases"]
        self.relevant_meta_tiers = [[] for _ in self.meta_tiers]
        self.var_check = IntVar()
        self.xl_header = ['id', 'tier type', 'annotation value', 'file', 'video_file', 'id1', 'id2',
                          'ts_start', 'ts_end', 'time_start', 'time_end', 'milli_start', 'milli_end',
                          "context", "phases", "subphases"]
        self.vid = None
        self.photo = None
        self.selected_dir = None
        self._job = None

        self.bt_width = 20
        self.cb_width = 30
        self.delay = 5
        self.max_fps = 0
        self.prev_timestamp = 0


        # create a label, entrybox and button for 'ELAN folder'
        Label(self.root, text="Choose ELAN folder", background='light grey').grid(row=1, column=1)
        self.entry_Elan_folder = Entry(self.root, width=50)
        self.entry_Elan_folder.grid(row=2, column=1, padx=30)
        self.entry_Elan_folder.insert(0, "")
        self.bt_Elan_folder = Button(self.root, text="Browse", width=self.bt_width, command=lambda: self.ask_elan_directory())
        self.bt_Elan_folder.grid(row=3, column=1)

        self.progress = ttk.Progressbar(self.root, orient=HORIZONTAL,
                                        length=self.bt_width*7.5, mode='determinate')
        self.progress.grid(row=4, column=1)
        self.label_files = Label(text="0 file uploaded", background='light grey')
        self.label_files.grid(row=5, column=1, pady=5)
        # Label(text="", background='light grey').grid(row=6, column=1)

        # create a label and combobox 'linguistic_type'
        Label(text="Select a tier type", background='light grey').grid(row=7, column=1)
        self.combo_linguistic_type_ref = ttk.Combobox(self.root, textvariable=StringVar(), width=self.cb_width)
        self.combo_linguistic_type_ref.grid(row=8, column=1)
        self.combo_linguistic_type_ref.bind("<<ComboboxSelected>>", self.combo_linguistic_type_ref_update)

        # create a label and combobox 'annotation value'
        Label(text="Select an annotation value", background='light grey').grid(row=9, column=1)
        self.combo_annotation_value = ttk.Combobox(self.root, textvariable=StringVar(), width=self.cb_width)
        self.combo_annotation_value.grid(row=10, column=1)
        self.combo_annotation_value.bind("<<ComboboxSelected>>", self.combo_annotation_value_update)

        # Label(text="", background='light grey').grid(row=11, column=1)
        # self.check_filter = Checkbutton(self.root, text='Filters On/Off', variable=self.var_check,
        #                                 background='light grey', onvalue=1, offvalue=0, command=self.filter_on_off)
        # self.check_filter.grid(row=12, column=1)

        # create a label and combobox 'filter'
        Label(text="Context filter", background='light grey').grid(row=13, column=1)
        self.combo_ctx = ttk.Combobox(self.root, textvariable=StringVar(), width=self.cb_width)
        self.combo_ctx.grid(row=14, column=1)
        self.combo_ctx.bind("<<ComboboxSelected>>", self.change_filter)
        Label(text="Phase filter", background='light grey').grid(row=15, column=1)
        self.combo_phs = ttk.Combobox(self.root, textvariable=StringVar(), width=self.cb_width)
        self.combo_phs.grid(row=16, column=1)
        self.combo_phs.bind("<<ComboboxSelected>>", self.change_filter)
        Label(text="Subphase filter", background='light grey').grid(row=17, column=1)
        self.combo_sph = ttk.Combobox(self.root, textvariable=StringVar(), width=self.cb_width)
        self.combo_sph.grid(row=18, column=1)
        self.combo_sph.bind("<<ComboboxSelected>>", self.change_filter)

        self.bt_reset_filters = Button(self.root, text="Reset filters", width=self.bt_width, command=lambda: self.reset_filter())
        self.bt_reset_filters.grid(row=19, column=1)

        self.label_files = Label(text="0 file uploaded", background='light grey')
        self.label_files.grid(row=5, column=1, pady=5)

        self.label_found = Label(text="0 result found", background='light grey')
        self.label_found.grid(row=21, column=1)

        # create tree and vertical scrollbar
        self.tree = ttk.Treeview(self.root)
        self.vsb = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        self.vsb.place(x=633, y=335, height=225)
        self.tree.configure(yscrollcommand=self.vsb.set)
        self.tree['columns'] = ("Event", "File", "Context", "Phase", "Subphase", "Time")
        self.tree.column('#0', width=0, stretch=NO)
        self.tree.column("Event", width=50, anchor=CENTER)
        self.tree.column("File", width=250, anchor=W)
        self.tree.column("Context", width=75, anchor=CENTER)
        self.tree.column("Phase", width=75, anchor=CENTER)
        self.tree.column("Subphase", width=75, anchor=CENTER)
        self.tree.column("Time", width=75, anchor=CENTER)
        self.tree.heading("Event", text="Event")
        self.tree.heading("File", text="File", anchor=W)
        self.tree.heading("Context", text="Context")
        self.tree.heading("Phase", text="Phase")
        self.tree.heading("Subphase", text="Subphase")
        self.tree.heading("Time", text="Time")
        self.tree.grid(row=1, column=2, rowspan=10, columnspan=3,  padx=30, pady=5)
        self.tree.bind("<Double-1>", self.on_double_click)

        self.bt_exp_all = Button(self.root, text="Export All", width=self.bt_width, command=lambda: self.export_xl("all"))
        self.bt_exp_all.grid(row=11, column=2)
        self.bt_exp_selection = Button(self.root, text="Export Selection", width=self.bt_width, command=lambda: self.export_xl("selection"))
        self.bt_exp_selection.grid(row=11, column=3)

        # create canvas for video
        self.canvas = Canvas(self.root, width=600, height=360)
        self.canvas.grid(row=12, column=2, rowspan=9, columnspan=3, pady=5)

        self.bt_open_in_elan = Button(self.root, text="Open in ELAN", width=self.bt_width, command=lambda: self.open_in_elan())
        self.bt_open_in_elan.grid(row=21, column=2, pady=5)
        self.bt_stop_playing = Button(self.root, text="Stop playing", width=self.bt_width, command=lambda: self.stop_playing())
        self.bt_stop_playing.grid(row=21, column=3, pady=5)

        self.slider = ttk.Scale(self.root, from_=1, to=60, value=25, orient='horizontal')
        self.slider.grid(row=21, column=4)

        # loop
        self.root.mainloop()


    def export_xl(self, span):  # xlsxwriter library stores data to excel
        now = datetime.now().strftime("%Y%m%d%H%M%S")
        wb = xw.Workbook(os.path.join(os.getcwd(), "output " + now + ".xlsx"))  # Create workbook
        ws = wb.add_worksheet("elan_data")  # Create child table
        ws.activate()  # Activation table
        ws.write_row('A1', self.xl_header)  # Write the header from cell A1
        i = 2  # Start writing data from the second line
        data = []

        if span == "all":
            full_dico = self.annotations
            for tier_type in full_dico:
                for annotation_value in full_dico[tier_type]:
                    dico = full_dico[tier_type][annotation_value]
                    for key in dico:
                        data.append(key)
                        for sub_key in dico[key]:
                            data.append(dico[key][sub_key])
                        ws.write_row('A' + str(i), data)
                        data = []
                        i += 1

        elif span == "selection":
            dico = self.annotations[self.combo_linguistic_type_ref.get()][self.combo_annotation_value.get()]
            for key in dico:
                data.append(key)
                for sub_key in dico[key]:
                    data.append(dico[key][sub_key])
                ws.write_row('A' + str(i), data)
                data = []
                i += 1
        # print(json.dumps(dico, sort_keys=False, indent=4))
        wb.close()  # Close table

    def stop_playing(self):
        if self._job is not None:
            self.root.after_cancel(self._job)
            self._job = None

    def open_in_elan(self):
        selected_item = self.tree.focus()
        clip_key = self.tree.item(selected_item)['values'][0]
        filename = self.annotations[self.combo_linguistic_type_ref.get()][self.combo_annotation_value.get()][
                                    str(clip_key)]["file"]
        path = os.path.join(self.selected_dir, filename)
        if sys.platform == "win32":
            os.startfile(path)
        else:
            opener = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.call([opener, path])

    def change_filter(self, event):
        self.update_tree()

    def filter_on_off(self):
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.update()

    def load_eaf_lists(self):
        self.reset_search()
        j = 0
        done_file = 0

        num_of_files = len([f for f in os.listdir(self.selected_dir) if f.endswith('.eaf')])
        percentage = 100 / num_of_files

        for file in os.listdir(self.selected_dir):

            if file.endswith(".eaf"):

                filename = os.path.join(self.selected_dir, file)

                tree = ET.parse(filename)
                root = tree.getroot()
                root = ET.tostring(root, encoding='unicode', method='xml')
                xmlstr = ET.fromstring(root)

                for header in xmlstr.findall('HEADER'):
                    for descriptor in header.findall('MEDIA_DESCRIPTOR'):
                        video_path = descriptor.attrib['MEDIA_URL']
                        seperator = max(video_path.rfind('/'), video_path.rfind(r"'\'"))
                        video_file = video_path[-(len(video_path) - seperator - 1):]

                # time_slots dict
                for time_order in xmlstr.findall('TIME_ORDER'):
                    for time_slot in time_order.findall('TIME_SLOT'):
                        self.time_slots[time_slot.attrib['TIME_SLOT_ID']] = {}
                        self.time_slots[time_slot.attrib['TIME_SLOT_ID']]['time'] = int(time_slot.attrib['TIME_VALUE'])
                        # for meta in self.meta_tiers:
                        #     self.time_slots[time_slot.attrib['TIME_SLOT_ID']][meta.lower()] = ""
                self.reverse_time_slots()
                # self.complete_time_slots()

                for tier in xmlstr.findall('TIER'):

                    if tier.attrib['TIER_ID'] == "ID_1":
                        for annotation in tier.findall('ANNOTATION'):
                            for align_annotation in annotation.findall('ALIGNABLE_ANNOTATION'):
                                for annotation_value in align_annotation.findall('ANNOTATION_VALUE'):
                                    id1 = annotation_value.text
                    if tier.attrib['TIER_ID'] == "ID-2":
                        for annotation in tier.findall('ANNOTATION'):
                            for align_annotation in annotation.findall('ALIGNABLE_ANNOTATION'):
                                for annotation_value in align_annotation.findall('ANNOTATION_VALUE'):
                                    id2 = annotation_value.text

                    if tier.attrib['LINGUISTIC_TYPE_REF'] in self.meta_tiers:
                        k = self.meta_tiers.index(tier.attrib['LINGUISTIC_TYPE_REF'])
                        for annotation in tier.findall('ANNOTATION'):
                            for align_annotation in annotation.findall('ALIGNABLE_ANNOTATION'):
                                ts_start_meta = self.get_ts_int(align_annotation.attrib['TIME_SLOT_REF1'], 2)
                                ts_end_meta = self.get_ts_int(align_annotation.attrib['TIME_SLOT_REF2'], 2)
                                for annotation_value in align_annotation.findall('ANNOTATION_VALUE'):
                                    for i in range(int(ts_start_meta), int(ts_end_meta) + 1):
                                        time_slot = "ts" + str(i)
                                        self.reversed_time_slots[self.time_slots[time_slot]['time']][
                                            tier.attrib['LINGUISTIC_TYPE_REF'].lower()] = annotation_value.text
                                    if annotation_value.text in self.relevant_meta_tiers[k]:
                                        pass
                                    else:
                                        self.relevant_meta_tiers[k].append(annotation_value.text)

                    if not tier.attrib['LINGUISTIC_TYPE_REF'] in self.all_tiers:
                        self.all_tiers.append(tier.attrib['LINGUISTIC_TYPE_REF'])
                        self.relevant_annotation_values.append([])

                    i = self.all_tiers.index(tier.attrib['LINGUISTIC_TYPE_REF'])
                    if not self.keys_exists(self.annotations, tier.attrib['LINGUISTIC_TYPE_REF']):
                        self.annotations[tier.attrib['LINGUISTIC_TYPE_REF']] = {}
                    for annotation in tier.findall('ANNOTATION'):
                        for align_annotation in annotation.findall('ALIGNABLE_ANNOTATION'):
                            ts_start, ts_end = self.get_time(align_annotation.attrib['TIME_SLOT_REF1'],
                                                             align_annotation.attrib['TIME_SLOT_REF2'], "hh:mm:ss")
                            # print(' |-->', align_annotation.attrib['ANNOTATION_ID'])
                            for annotation_value in align_annotation.findall('ANNOTATION_VALUE'):
                                if not self.keys_exists(self.annotations, tier.attrib['LINGUISTIC_TYPE_REF'],
                                                        annotation_value.text):
                                    self.annotations[tier.attrib['LINGUISTIC_TYPE_REF']][annotation_value.text] = {}
                                    # print("adding: " + tier.attrib['LINGUISTIC_TYPE_REF'] + " " + annotation_value.text)
                                else:
                                    pass
                                    # print("already added " + tier.attrib['LINGUISTIC_TYPE_REF'] + " " + annotation_value.text)

                                self.annotations[tier.attrib['LINGUISTIC_TYPE_REF']][annotation_value.text][str(j)] = {}
                                new_annot = self.annotations[tier.attrib['LINGUISTIC_TYPE_REF']][annotation_value.text][
                                    str(j)]

                                new_annot["tier type"] = tier.attrib['LINGUISTIC_TYPE_REF']
                                new_annot["annotation value"] = annotation_value.text
                                new_annot["file"] = file
                                try:
                                    new_annot["video_file"] = video_file
                                except:
                                    pass

                                new_annot["id1"] = None
                                new_annot["id2"] = None
                                try:
                                    new_annot["id1"] = id1
                                    new_annot["id2"] = id2
                                except:
                                    pass

                                new_annot["ts_start"] = align_annotation.attrib['TIME_SLOT_REF1']
                                new_annot["ts_end"] = align_annotation.attrib['TIME_SLOT_REF2']
                                new_annot["time_start"] = ts_start
                                new_annot["time_end"] = ts_end
                                new_annot["milli_start"] = self.time_slots[align_annotation.attrib['TIME_SLOT_REF1']][
                                    'time']
                                new_annot["milli_end"] = self.time_slots[align_annotation.attrib['TIME_SLOT_REF2']][
                                    'time']

                                for meta in self.meta_tiers:
                                    new_annot[meta.lower()] = self.reversed_time_slots[new_annot["milli_start"]][
                                        meta.lower()]

                                if annotation_value.text in (item for sublist in self.relevant_annotation_values for
                                                             item in sublist):
                                    pass
                                else:
                                    self.relevant_annotation_values[i].append(annotation_value.text)
                                # print(' |---->', annotation_value.text)

                                j += 1

                done_file += 1
                # print(json.dumps(self.reversed_time_slots, sort_keys=False, indent=4))
                self.bar("update", num_of_files, done_file, percentage)
        self.bar("finished", num_of_files, 0, 0)
        # self.export_all(self.annotations["Behaviour"]["Approach"])

        # print(json.dumps(self.annotations, sort_keys=False, indent=4))
        # print(json.dumps(self.time_slots, sort_keys=False, indent=4))
        # print(len(self.all_tiers))
        # print(len(self.relevant_annotation_values))
        # print(self.relevant_meta_tiers)

    def reset_filter(self):
        self.combo_ctx.set('')
        self.combo_phs.set('')
        self.combo_sph.set('')
        self.update_tree()

    def reset_search(self):
        self.time_slots = {}
        self.all_tiers = []
        self.annotations = {}
        self.annotation_values = []
        self.relevant_annotation_values = []
        self.combo_linguistic_type_ref.set('')
        self.combo_annotation_value.set('')
        self.entry_Elan_folder.delete(0, 'end')
        self.tree.delete(*self.tree.get_children())

    def reverse_time_slots(self):
        for time_slot in self.time_slots:
            if not self.keys_exists(self.reversed_time_slots, self.time_slots[time_slot]['time']):
                self.reversed_time_slots[self.time_slots[time_slot]['time']] = {}
            for meta in self.meta_tiers:
                self.reversed_time_slots[self.time_slots[time_slot]['time']][meta.lower()] = ""
        # print(json.dumps(self.reversed_time_slots, sort_keys=False, indent=4))

    def get_ts_int(self, string, num_char):
        right_str = string[-(len(string) - num_char):]
        return right_str

    def get_meta(self, ts_str):
        ts_int = ts_str[-(len(ts_str) - 2):]
        return in_context, in_phase, in_subphase

    def keys_exists(self, element, *keys):
        '''
        Check if *keys (nested) exists in `element` (dict).
        '''
        if not isinstance(element, dict):
            raise AttributeError('keys_exists() expects dict as first argument.')
        if len(keys) == 0:
            raise AttributeError('keys_exists() expects at least two arguments, one given.')

        _element = element
        for key in keys:
            try:
                _element = _element[key]
            except KeyError:
                return False
        return True

    def bar(self, state, num_of_files, done_file, percentage):
        if state == "update":
            self.progress['value'] = done_file * percentage
            self.label_files.config(text='Analyzing %d out of %s files. Please wait...' % (done_file, num_of_files))
        elif state == "finished":
            self.label_files.config(text='Done! %s files analyzed!' % num_of_files)
        self.root.update_idletasks()

    def on_double_click(self, event):
        selected_item = self.tree.focus()
        clip_key = self.tree.item(selected_item)['values'][0]
        video_name = self.tree.item(selected_item)['values'][1]

        start_milli = \
            self.annotations[self.combo_linguistic_type_ref.get()][self.combo_annotation_value.get()][str(clip_key)][
                "milli_start"]

        self.vid = MyVideoCapture(start_milli, os.path.join(self.selected_dir, video_name))
        self.stop_playing()
        self.prev_timestamp = time.time()
        self.update_frame()

    def update_frame(self):
        # Get a frame from the video source
        time_elapsed = time.time() - self.prev_timestamp
        if time_elapsed > 1. / self.slider.get():
            ret, frame = self.vid.get_frame()
            self.prev_timestamp = time.time()
            if ret:
                self.photo = PIL.ImageTk.PhotoImage(image=PIL.Image.fromarray(frame))
                self.canvas.create_image(0, 0, image=self.photo, anchor=NW)
                self._job = self.root.after(self.delay, self.update_frame)
        else:
            self._job = self.root.after(self.delay, self.update_frame)

    def update_tree(self):
        i = 0
        self.tree.delete(*self.tree.get_children())
        dico = self.annotations[self.combo_linguistic_type_ref.get()][self.combo_annotation_value.get()]
        for record in dico:
            bl_ctx, bl_phs, bl_sph = True, True, True
            if len(self.combo_ctx.get()):
                if not dico[record]['context'] == self.combo_ctx.get():
                    bl_ctx = False
            if len(self.combo_phs.get()):
                if not dico[record]['phases'] == self.combo_phs.get():
                    bl_phs = False
            if len(self.combo_sph.get()):
                if not dico[record]['subphases'] == self.combo_sph.get():
                    bl_sph = False
            if bl_ctx and bl_phs and bl_sph:
                i += 1
                self.tree.insert(parent='', index='end', iid=None, text="",
                                 values=(
                                 record, dico[record]['video_file'], dico[record]['context'], dico[record]['phases'],
                                 dico[record]['subphases'], dico[record]['time_start']))

        if i > 1:
            self.label_found.config(text='%s results found' % i)
        else:
            self.label_found.config(text='%s result found' % i)

    def ask_elan_directory(self):
        self.selected_dir = filedialog.askdirectory()
        self.load_eaf_lists()
        self.entry_Elan_folder.insert(0, self.selected_dir)
        self.combo_linguistic_type_ref['values'] = self.all_tiers
        self.combo_ctx['values'] = self.relevant_meta_tiers[0]
        self.combo_phs['values'] = self.relevant_meta_tiers[1]
        self.combo_sph['values'] = self.relevant_meta_tiers[2]

    def combo_linguistic_type_ref_update(self, event):
        ind = self.all_tiers.index(self.combo_linguistic_type_ref.get())
        self.combo_annotation_value['values'] = tuple(self.relevant_annotation_values[ind])
        self.combo_annotation_value.set('')
        self.tree.delete(*self.tree.get_children())
        self.label_found.config(text='%s result found' % 0)

    def combo_annotation_value_update(self, event):
        self.update_tree()
        return self.annotation_values

    def combo_filter_update(self, event):
        pass

    def get_time(self, ts_start, ts_end, format_style):
        millis_start = self.time_slots[ts_start]['time']
        millis_end = self.time_slots[ts_end]['time']

        if format_style == "hh:mm:ss":
            time_start = self.conv_millis_to_hh_mm_ss(millis_start)
            time_end = self.conv_millis_to_hh_mm_ss(millis_end)
        else:
            time_start = millis_start
            time_end = millis_end

        return time_start, time_end

    def conv_millis_to_hh_mm_ss(self, millis):
        millis = int(millis)
        seconds = (millis / 1000) % 60
        seconds = int(seconds)
        seconds = str(seconds).zfill(2)
        minutes = (millis / (1000 * 60)) % 60
        minutes = int(minutes)
        minutes = str(minutes).zfill(2)
        hours = (millis / (1000 * 60 * 60)) % 24
        hours = int(hours)
        hours = str(hours).zfill(2)

        return ':'.join([str(hours), str(minutes), str(seconds)])


class MyVideoCapture:
    def __init__(self, start_milli=0, video_source=0):
        # Open the video source
        self.vid = cv2.VideoCapture(video_source)
        if not self.vid.isOpened():
            raise ValueError("Unable to open video source", video_source)

        self.vid.set(0, int(start_milli))

        # Get video source width and height
        self.width = self.vid.get(cv2.CAP_PROP_FRAME_WIDTH)
        self.height = self.vid.get(cv2.CAP_PROP_FRAME_HEIGHT)
        self.dim = self.get_dim(self.width, self.height)

    def get_frame(self):
        while self.vid.isOpened():

            ret, frame = self.vid.read()

            if ret:
                # Return a boolean success flag and the current frame converted to BGR
                return ret, cv2.resize(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB), self.dim)
            else:
                return ret, None
        else:
            return False, None

    def get_dim(self, width, height):
        if width == 1920 and height == 1080:
            divider = 3
        elif width == 1080 and height == 720:
            divider = 2
        else:
            divider = 4
        return int(width / divider), int(height / divider)

    # Release the video source when the object is destroyed
    def __del__(self):
        if self.vid.isOpened():
            self.vid.release()


if __name__ == "__main__":
    app = ElanFinder()
