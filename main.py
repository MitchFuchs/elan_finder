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
import pandas as pd
from datetime import datetime
import time
# import json

class ElanFinder:
    def __init__(self):
        self._col = ['linguistic_type_ref', 'tier_id', 'annotation_value', 'annotation_id', 'time_slot_ref1',
                     'time_slot_ref2']
        self._head = ['ANNOTATION_DOCUMENT', 'HEADER', 'MEDIA_DESCRIPTOR', 'PROPERTY']
        self._time = ['TIME_ORDER', 'TIME_SLOT']
        self._time_attrib = ['TIME_VALUE', 'TIME_SLOT_ID']
        self._annot = ['TIER', 'ANNOTATION', 'ALIGNABLE_ANNOTATION', 'ANNOTATION_VALUE']
        self._tier = ['LINGUISTIC_TYPE_REF', 'TIER_ID']
        self._align = ['ANNOTATION_ID', 'CVE_REF', 'TIME_SLOT_REF1', 'TIME_SLOT_REF2']
        self._cps = ['Context', 'Phases', 'Subphases']
        self._cps_ts = self._cps + [self._time_attrib[1]]
        self._cps_ts1 = ['Context_ts1', 'Phases_ts1', 'Subphases_ts1', 'Time_value_ts1']
        self._cps_ts2 = ['Context_ts2', 'Phases_ts2', 'Subphases_ts2', 'Time_value_ts2']

        self._heading_tree = ["Event", "File", "Context", "Phase", "Subphase", "Time"]
        self._col_tree = ["index", "file", "Context_ts1", "Phases_ts1", "Subphases_ts1", "Time_milli_ts1"]
        self._header = ['AUTHOR', 'DATE', 'FORMAT', 'VERSION', 'MEDIA_FILE', 'TIME_UNITS', 'MEDIA_URL', 'MIME_TYPE']
        self._path = ['selected_dir', 'file', 'media']
        self.selected_dirs = []

        self.df = pd.DataFrame()

        self.vid = None
        self.photo = None
        self.selected_dir = None
        self._job = None

        self.bt_width = 20
        self.cb_width = 30
        self.delay = 5
        self.max_fps = 0
        self.prev_timestamp = 0

        self.root = Tk()
        self.root.configure(background='light grey')
        self.root.title("ELAN finder")

        # create a label, entrybox and button for 'ELAN folder'
        Label(self.root, text="Choose ELAN folder", background='light grey').grid(row=1, column=1)
        self.entry_Elan_folder = Entry(self.root, width=50)
        self.entry_Elan_folder.grid(row=2, column=1, padx=30)
        self.entry_Elan_folder.insert(0, "")
        self.bt_Elan_folder = Button(self.root, text="Browse", width=self.bt_width,
                                     command=lambda: self.ask_elan_directory("<Button-1>"))
        self.bt_Elan_folder.grid(row=3, column=1)

        self.progress = ttk.Progressbar(self.root, orient=HORIZONTAL,
                                        length=self.bt_width * 7.5, mode='determinate')
        self.progress.grid(row=4, column=1)
        self.label_files = Label(self.root, text="0 file uploaded", background='light grey')
        self.label_files.grid(row=5, column=1, pady=5)
        # Label(text="", background='light grey').grid(row=6, column=1)

        # create a label and combobox 'linguistic_type'
        Label(self.root, text="Select a tier type", background='light grey').grid(row=7, column=1)
        self.combo_linguistic_type_ref = ttk.Combobox(self.root, textvariable=StringVar(), width=self.cb_width)
        self.combo_linguistic_type_ref.grid(row=8, column=1)
        self.combo_linguistic_type_ref.bind("<<ComboboxSelected>>", self.combo_linguistic_type_ref_update)

        # create a label and combobox 'annotation value'
        Label(self.root, text="Select an annotation value", background='light grey').grid(row=9, column=1)
        self.combo_annotation_value = ttk.Combobox(self.root, textvariable=StringVar(), width=self.cb_width)
        self.combo_annotation_value.grid(row=10, column=1)
        self.combo_annotation_value.bind("<<ComboboxSelected>>", self.update_tree)

        # create a label and combobox 'filter'
        Label(self.root, text="Context filter", background='light grey').grid(row=13, column=1)
        self.combo_ctx = ttk.Combobox(self.root, textvariable=StringVar(), width=self.cb_width)
        self.combo_ctx.grid(row=14, column=1)
        self.combo_ctx.bind("<<ComboboxSelected>>", self.update_tree)
        Label(self.root, text="Phase filter", background='light grey').grid(row=15, column=1)
        self.combo_phs = ttk.Combobox(self.root, textvariable=StringVar(), width=self.cb_width)
        self.combo_phs.grid(row=16, column=1)
        self.combo_phs.bind("<<ComboboxSelected>>", self.update_tree)
        Label(self.root, text="Subphase filter", background='light grey').grid(row=17, column=1)
        self.combo_sph = ttk.Combobox(self.root, textvariable=StringVar(), width=self.cb_width)
        self.combo_sph.grid(row=18, column=1)
        self.combo_sph.bind("<<ComboboxSelected>>", self.update_tree)

        # self.bt_reset_filters = Button(self.root, text="Reset filters", width=self.bt_width,
        #                                command=lambda: self.reset_filter())
        self.bt_reset_filters = Button(self.root, text="Reset filters", width=self.bt_width)
        self.bt_reset_filters.bind("<Button-1>", self.reset_filter)
        self.bt_reset_filters.grid(row=19, column=1, pady=5)

        Label(self.root, text="", background='light grey', height=12).grid(row=20, column=1)

        self.label_files = Label(self.root, text="0 file uploaded", background='light grey')
        self.label_files.grid(row=5, column=1, pady=5)

        self.label_found = Label(self.root, text="0 result found", background='light grey')
        self.label_found.grid(row=21, column=1)

        # create tree and vertical scrollbar
        self.tree = ttk.Treeview(self.root)
        self.vsb = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        self.vsb.place(x=633, y=335, height=225)
        self.tree.configure(yscrollcommand=self.vsb.set)
        self.tree['columns'] = ("Event", "File", "Context", "Phase", "Subphase", "Time")
        self.tree.column('#0', width=0, stretch=NO)

        self.tree.column(self._heading_tree[0], width=50, anchor=CENTER)
        self.tree.column(self._heading_tree[1], width=250, anchor=W)
        self.tree.column(self._heading_tree[2], width=75, anchor=CENTER)
        self.tree.column(self._heading_tree[3], width=75, anchor=CENTER)
        self.tree.column(self._heading_tree[4], width=75, anchor=CENTER)
        self.tree.column(self._heading_tree[5], width=75, anchor=CENTER)
        self.tree.heading(self._heading_tree[0], text=self._heading_tree[0])
        self.tree.heading(self._heading_tree[1], text=self._heading_tree[1], anchor=W)
        self.tree.heading(self._heading_tree[2], text=self._heading_tree[2])
        self.tree.heading(self._heading_tree[3], text=self._heading_tree[3])
        self.tree.heading(self._heading_tree[4], text=self._heading_tree[4])
        self.tree.heading(self._heading_tree[5], text=self._heading_tree[5])
        self.tree.grid(row=1, column=2, rowspan=10, columnspan=3, padx=30, pady=5)

        self.tree.bind("<Double-1>", self.on_double_click)

        self.bt_exp_all = Button(self.root, text="Export All", width=self.bt_width,
                                 command=lambda: self.export('all'))
        self.bt_exp_all.grid(row=11, column=2)
        self.bt_exp_selection = Button(self.root, text="Export Selection", width=self.bt_width,
                                       command=lambda: self.export("selection"))
        self.bt_exp_selection.grid(row=11, column=3)

        # create canvas for video
        self.canvas = Canvas(self.root, width=600, height=360)
        self.canvas.grid(row=12, column=2, rowspan=9, columnspan=3, pady=5)

        self.bt_open_in_elan = Button(self.root, text="Open in ELAN", width=self.bt_width,
                                      command=lambda: self.open_in_elan())
        self.bt_open_in_elan.grid(row=21, column=2, pady=5)
        self.bt_stop_playing = Button(self.root, text="Stop playing", width=self.bt_width,
                                      command=lambda: self.stop_playing())
        self.bt_stop_playing.grid(row=21, column=3, pady=5)

        self.slider = ttk.Scale(self.root, from_=1, to=60, value=25, orient='horizontal')
        self.slider.grid(row=21, column=4)

        # loop
        self.root.mainloop()

    def ask_elan_directory(self, event):
        self.selected_dir = filedialog.askdirectory()
        if self.selected_dir not in self.selected_dirs:
            self.selected_dirs.insert(0, self.selected_dir)
            self.df = pd.concat([self.df, self.load_files()], ignore_index=True)
            self.df['index'] = self.df.index
            self.entry_Elan_folder.insert(0, self.selected_dirs)
            self.update_tree(event)
            self.combo_linguistic_type_ref['values'] = self.df['linguistic_type_ref'].unique().tolist()
            self.combo_ctx['values'] = self.df['Context_ts1'].dropna().unique().tolist()
            self.combo_phs['values'] = self.df['Phases_ts1'].dropna().unique().tolist()
            self.combo_sph['values'] = self.df['Subphases_ts1'].dropna().unique().tolist()

    def load_files(self):
        self.reset_search()
        data = []
        done_file = 0
        files = [f for f in os.listdir(self.selected_dir) if f.endswith('.eaf')]
        percentage = 100 / len(files)
        for file in files:
            data.extend(self.parse_eaf_file(file))
            done_file += 1
            self.bar("update", len(files), done_file, percentage)
        cols = self._col + self._cps_ts1 + self._cps_ts2 + self._header + self._path
        df = pd.DataFrame(data=data, columns=cols)
        df[['Time_milli_ts1', 'Time_milli_ts2']] = df.loc[:,['Time_value_ts1', 'Time_value_ts2']].applymap(self.conv_millis_to_hh_mm_ss)
        self.bar("finished", len(files), 0, 0)
        return df

    def parse_eaf_file(self, file):
        header = {}
        time_slots = []
        values = []
        xroot = ET.parse(os.path.join(self.selected_dir, file)).getroot()
        for descendant in xroot.iter():
            if descendant.tag in self._head:
                # general information from the header
                header.update(descendant.attrib)
            elif descendant.tag in self._time and descendant.attrib:
                # general information about time slots
                time_slots.append([descendant.attrib[self._time_attrib[1]], descendant.attrib[self._time_attrib[0]]])
            elif descendant.tag in self._annot:
                if descendant.tag == self._annot[0]:  # TIER
                    linguistic_type_ref, tier_id = descendant.attrib[self._tier[0]], descendant.attrib[self._tier[1]]
                elif descendant.tag == self._annot[2]:  # ALIGNABLE_ANNOTATION
                    annotation_id, ts_ref1, ts_ref2 = descendant.attrib[self._align[0]], descendant.attrib[self._align[2]], \
                                                      descendant.attrib[self._align[3]]
                    # cve_ref = descendant.attrib[_align[1]]
                elif descendant.tag == self._annot[3]:  # ANNOTATION_VALUE
                    if linguistic_type_ref in self._cps:
                        if type(time_slots) is list:
                            time_slots = pd.DataFrame.from_records(time_slots, columns=self._time_attrib)
                            time_slots['ts'] = time_slots[self._time_attrib[0]].str.split(pat='ts', expand=True)[1].astype(int)
                        mask_ts = (time_slots['ts'].ge(int(ts_ref1.split('ts')[1])) & time_slots['ts'].le(
                            int(ts_ref2.split('ts')[1])))
                        time_slots.loc[mask_ts, linguistic_type_ref] = descendant.text
                    else:
                        values.append([linguistic_type_ref, tier_id, descendant.text, annotation_id, ts_ref1, ts_ref2])
        for value in values:
            value.extend(time_slots.loc[time_slots[self._time_attrib[0]] == value[4], self._cps_ts].iloc[0].tolist())
            value.extend(time_slots.loc[time_slots[self._time_attrib[0]] == value[5], self._cps_ts].iloc[0].tolist())
            value.extend([header[x] for x in self._header])
            value.extend([self.selected_dir, file])
            value.append(self.media_name(header))
        return values

    def media_name(self, h):
        media = h['MEDIA_URL']
        slash = media.rfind('/') + 1
        backslash = media.rfind('\\') + 1
        return media[max(slash, backslash):]

    def mask_from_filters(self):
        mask = (self.df['linguistic_type_ref'] == self.combo_linguistic_type_ref.get()) & (
                self.df['annotation_value'] == self.combo_annotation_value.get())
        if self.combo_ctx.get(): mask = mask & (self.df['Context_ts1'] == self.combo_ctx.get())
        if self.combo_phs.get(): mask = mask & (self.df['Phases_ts1'] == self.combo_phs.get())
        if self.combo_sph.get(): mask = mask & (self.df['Subphases_ts1'] == self.combo_sph.get())

        return mask

    def update_tree(self, event):
        self.tree.delete(*self.tree.get_children())
        mask = self.mask_from_filters()
        for index, row in self.df[mask].iterrows():
            self.tree.insert("", 'end', text='', iid=None,
                             values=list(row.loc[self._col_tree].where(pd.notnull(row.loc[self._col_tree]), '')))
        num_rows = len(self.df[mask])
        if num_rows > 1:
            self.label_found.config(text='%s results found' % num_rows)
        else:
            self.label_found.config(text='%s result found' % num_rows)

    def combo_linguistic_type_ref_update(self, event):
        mask = self.df['linguistic_type_ref'] == self.combo_linguistic_type_ref.get()
        self.combo_annotation_value['values'] = self.df[mask]['annotation_value'].dropna().unique().tolist()
        self.combo_annotation_value.set('')
        self.tree.delete(*self.tree.get_children())
        self.label_found.config(text='%s result found' % 0)

    def on_double_click(self, event):
        clip_key = self.tree.item(self.tree.focus())['values'][0]
        start_milli = self.df.iloc[clip_key].loc['Time_value_ts1']
        media = self.df.iloc[clip_key].loc['media']
        in_dir = self.df.iloc[clip_key].loc['selected_dir']
        self.vid = MyVideoCapture(start_milli, os.path.join(in_dir, media))
        self.stop_playing()
        self.update_frame()

    def export(self, span):
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        if span == 'selection':
            self.df.loc[self.mask_from_filters()].to_excel(f'output {timestamp}.xlsx')
        else:
            self.df.to_excel(f'output {timestamp}.xlsx')

    def open_in_elan(self):
        clip_key = self.tree.item(self.tree.focus())['values'][0]
        eaf_file = self.df.iloc[clip_key].loc['file']
        in_dir = self.df.iloc[clip_key].loc['selected_dir']
        path = os.path.join(in_dir, eaf_file)
        if sys.platform == "win32":
            os.startfile(path)
        else:
            opener = "open" if sys.platform == "darwin" else "xdg-open"
            subprocess.call([opener, path])

    def reset_filter(self, event):
        self.combo_ctx.set('')
        self.combo_phs.set('')
        self.combo_sph.set('')
        self.update_tree(event)

    def reset_search(self):
        self.combo_linguistic_type_ref.set('')
        self.combo_annotation_value.set('')
        # self.entry_Elan_folder.delete(0, 'end')
        self.tree.delete(*self.tree.get_children())

    def bar(self, state, num_of_files, done_file, percentage):
        if state == "update":
            self.progress['value'] = done_file * percentage
            self.label_files.config(text='Analyzing %d out of %s files. Please wait...' % (done_file, num_of_files))
        elif state == "finished":
            self.label_files.config(text='Done! %s files analyzed!' % num_of_files)
        self.root.update_idletasks()

    def conv_millis_to_hh_mm_ss(self, millis):
        millis = int(millis)
        seconds = str(int((millis / 1000) % 60)).zfill(2)
        minutes = str(int((millis / (1000 * 60)) % 60)).zfill(2)
        hours = str(int((millis / (1000 * 60 * 60)) % 24)).zfill(2)

        return ':'.join([str(hours), str(minutes), str(seconds)])

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

    def stop_playing(self):
        if self._job is not None:
            self.root.after_cancel(self._job)
            self._job = None

class MyVideoCapture:
    def __init__(self, start_milli, video_source):
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
