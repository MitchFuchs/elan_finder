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
import os
import xml.etree.ElementTree as ET
import json

#test

class ElanFinder:

    def __init__(self):
        self.root = Tk()
        self.root.configure(background='light grey')
        self.root.title("ELAN finder")

        self.selected_dir = None
        self._job = None

        self.time_slots = {}

        # self.relevant_tiers = ["Gesture", "Behaviour"]
        self.all_tiers = []
        self.annotations = {}
        self.annotation_values = []
        self.relevant_annotation_values = []
        self.vid = None
        self.photo = None

        # create a label, entrybox and button for 'ELAN folder'
        Label(self.root, text="Choose ELAN folder", background='light grey').grid(row=1, column=1)
        self.entry_Elan_folder = Entry(self.root, width=50)
        self.entry_Elan_folder.grid(row=2, column=1)
        self.entry_Elan_folder.insert(0, "")
        self.button_Elan_folder = Button(self.root, text="Browse", width=42, command=lambda: self.ask_elan_directory())
        self.button_Elan_folder.grid(row=3, column=1)

        self.progress = ttk.Progressbar(self.root, orient=HORIZONTAL,
                               length=300, mode='determinate')
        self.progress.grid(row=4, column=1)
        # self.progress['value'] = 100
        # Label(self.root, text="Choose ELAN folder").grid(row=4, column=1)
        self.label_files = Label(text="0 file uploaded", background='light grey')
        self.label_files.grid(row=5, column=1)
        Label(text="", background='light grey').grid(row=6, column=1)

        # create a label and combobox 'linguistic_type'
        Label(text="Select a linguistic type", background='light grey').grid(row=7, column=1)
        self.combo_linguistic_type_ref = ttk.Combobox(self.root, textvariable=StringVar(), width=48)
        self.combo_linguistic_type_ref.grid(row=8, column=1)
        self.combo_linguistic_type_ref.bind("<<ComboboxSelected>>", self.combo_linguistic_type_ref_update)

        # create a label and combobox 'annotation value'
        Label(text="Select an annotation value", background='light grey').grid(row=9, column=1)
        self.combo_annotation_value = ttk.Combobox(self.root, textvariable=StringVar(), width=48)
        self.combo_annotation_value.grid(row=10, column=1)
        self.combo_annotation_value.bind("<<ComboboxSelected>>", self.combo_annotation_value_update)

        # create tree and vertical scrollbar
        self.tree = ttk.Treeview(self.root)
        self.vsb = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        self.vsb.place(x=608, y=245, height=225)
        self.tree.configure(yscrollcommand=self.vsb.set)
        self.tree['columns'] = ("Clip", "File", "ID1", "ID2", "Time")
        self.tree.column('#0', width=0, stretch=NO)
        self.tree.column("Clip", width=50, anchor=CENTER)
        self.tree.column("File", width=300, anchor=W)
        self.tree.column("ID1", width=75, anchor=CENTER)
        self.tree.column("ID2", width=75, anchor=CENTER)
        self.tree.column("Time", width=75, anchor=CENTER)
        self.tree.heading("Clip", text="Clip")
        self.tree.heading("File", text="File", anchor=W)
        self.tree.heading("ID1", text="ID1")
        self.tree.heading("ID2", text="ID2")
        self.tree.heading("Time", text="Time")
        self.tree.grid(row=11, column=1, padx=30, pady=30)
        self.tree.bind("<Double-1>", self.on_double_click)

        # create canvas for video
        self.canvas = Canvas(self.root, width=600, height=360)
        self.canvas.grid(row=12, column=1)
        self.delay = 5

        # loop
        self.root.mainloop()

    def load_eaf_lists(self):
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
                        self.time_slots[time_slot.attrib['TIME_SLOT_ID']] = time_slot.attrib['TIME_VALUE']

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

                    if not tier.attrib['LINGUISTIC_TYPE_REF'] in self.all_tiers:
                    # print(type(self.all_tiers))
                    # if not self.keys_exists(self.all_tiers, tier.attrib['LINGUISTIC_TYPE_REF']):
                        # print(str(True) + ": " + tier.attrib['LINGUISTIC_TYPE_REF'])
                        self.all_tiers.append(tier.attrib['LINGUISTIC_TYPE_REF'])
                        self.relevant_annotation_values.append([])

                    # print("all tier: " + str(self.all_tiers))
                    # print("relevant annotation values: " + str(self.relevant_annotation_values))
                    # print(tier.attrib['LINGUISTIC_TYPE_REF'])
                    i = self.all_tiers.index(tier.attrib['LINGUISTIC_TYPE_REF'])
                    # print(i)
                    if not self.keys_exists(self.annotations, tier.attrib['LINGUISTIC_TYPE_REF']):
                        self.annotations[tier.attrib['LINGUISTIC_TYPE_REF']] = {}
                    # print(self.annotations)
                    for annotation in tier.findall('ANNOTATION'):
                        for align_annotation in annotation.findall('ALIGNABLE_ANNOTATION'):
                            ts_start, ts_end = self.get_time(align_annotation.attrib['TIME_SLOT_REF1'],
                                                             align_annotation.attrib['TIME_SLOT_REF2'], "hh:mm:ss")
                            # print(' |-->', align_annotation.attrib['ANNOTATION_ID'])
                            for annotation_value in align_annotation.findall('ANNOTATION_VALUE'):
                                if not self.keys_exists(self.annotations, tier.attrib['LINGUISTIC_TYPE_REF'], annotation_value.text):
                                    self.annotations[tier.attrib['LINGUISTIC_TYPE_REF']][annotation_value.text] = {}
                                    # print("adding: " + tier.attrib['LINGUISTIC_TYPE_REF'] + " " + annotation_value.text)
                                else:
                                    pass
                                    # print("already added " + tier.attrib['LINGUISTIC_TYPE_REF'] + " " + annotation_value.text)

                                self.annotations[tier.attrib['LINGUISTIC_TYPE_REF']][annotation_value.text][str(j)] = {}

                                self.annotations[tier.attrib['LINGUISTIC_TYPE_REF']][annotation_value.text][str(j)]["file"] = file
                                self.annotations[tier.attrib['LINGUISTIC_TYPE_REF']][annotation_value.text][str(j)][
                                    "video_file"] = video_file
                                self.annotations[tier.attrib['LINGUISTIC_TYPE_REF']][annotation_value.text][str(j)][
                                    "id1"] = "" #id1
                                self.annotations[tier.attrib['LINGUISTIC_TYPE_REF']][annotation_value.text][str(j)][
                                    "id2"] = "" #id2

                                self.annotations[tier.attrib['LINGUISTIC_TYPE_REF']][annotation_value.text][str(j)][
                                    "ts_start"] = align_annotation.attrib['TIME_SLOT_REF1']
                                self.annotations[tier.attrib['LINGUISTIC_TYPE_REF']][annotation_value.text][str(j)][
                                    "ts_end"] = align_annotation.attrib['TIME_SLOT_REF2']

                                self.annotations[tier.attrib['LINGUISTIC_TYPE_REF']][annotation_value.text][str(j)][
                                    "time_start"] = ts_start
                                self.annotations[tier.attrib['LINGUISTIC_TYPE_REF']][annotation_value.text][str(j)][
                                    "time_end"] = ts_end

                                self.annotations[tier.attrib['LINGUISTIC_TYPE_REF']][annotation_value.text][str(j)][
                                    "milli_start"] = self.time_slots[align_annotation.attrib['TIME_SLOT_REF1']]
                                self.annotations[tier.attrib['LINGUISTIC_TYPE_REF']][annotation_value.text][str(j)][
                                    "milli_end"] = self.time_slots[align_annotation.attrib['TIME_SLOT_REF2']]

                                if annotation_value.text in (item for sublist in self.relevant_annotation_values for
                                                             item in sublist):
                                    pass
                                else:
                                    self.relevant_annotation_values[i].append(annotation_value.text)
                                # print(' |---->', annotation_value.text)

                                j += 1

                done_file += 1
                self.bar("update", num_of_files, done_file, percentage)
        self.bar("finished", num_of_files, 0, 0)

        print(json.dumps(self.annotations, sort_keys=False, indent=4))
        # print(json.dumps(self.time_slots, sort_keys=False, indent=4))
        # print(len(self.all_tiers))
        # print(len(self.relevant_annotation_values))

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
            self.label_files.config(text='Done! %s files analyzed!' % (num_of_files))
        self.root.update_idletasks()

    def on_double_click(self, event):
        selected_item = self.tree.focus()
        clip_key = self.tree.item(selected_item)['values'][0]
        video_name = self.tree.item(selected_item)['values'][1]

        start_milli = \
            self.annotations[self.combo_linguistic_type_ref.get()][self.combo_annotation_value.get()][str(clip_key)][
                "milli_start"]

        self.vid = MyVideoCapture(start_milli, os.path.join(self.selected_dir, video_name))
        if self._job is not None:
            self.root.after_cancel(self._job)
            self._job = None
        self.update_frame()

    def update_frame(self):
        # Get a frame from the video source
        ret, frame = self.vid.get_frame()
        if ret:
            self.photo = PIL.ImageTk.PhotoImage(image=PIL.Image.fromarray(frame))
            self.canvas.create_image(0, 0, image=self.photo, anchor=NW)
            self._job = self.root.after(self.delay, self.update_frame)

    def update_tree(self):
        print("update tree")
        self.tree.delete(*self.tree.get_children())
        try:
            print(self.combo_linguistic_type_ref.get())
            print(self.combo_annotation_value.get())
            dico = self.annotations[self.combo_linguistic_type_ref.get()][self.combo_annotation_value.get()]
            print(dico)
            for record in dico:
                print(record)
                self.tree.insert(parent='', index='end', iid=None, text="",
                                 values=(record, dico[record]['video_file'], dico[record]['id1'], dico[record]['id2'],
                                         dico[record]['time_start']))
        except:
            pass

    def ask_elan_directory(self):
        self.selected_dir = filedialog.askdirectory()
        self.entry_Elan_folder.insert(0, self.selected_dir)
        self.load_eaf_lists()
        self.combo_linguistic_type_ref['values'] = self.all_tiers
        # prepare_label_linguistic_type()

    def combo_linguistic_type_ref_update(self, event):
        ind = self.all_tiers.index(self.combo_linguistic_type_ref.get())
        self.combo_annotation_value['values'] = tuple(self.relevant_annotation_values[ind])
        self.combo_annotation_value.set('')
        self.tree.delete(*self.tree.get_children())

    def combo_annotation_value_update(self, event):
        self.update_tree()
        return self.annotation_values

    def get_time(self, ts_start, ts_end, format_style):
        millis_start = self.time_slots[ts_start]
        millis_end = self.time_slots[ts_end]

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
        if self.vid.isOpened():
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
