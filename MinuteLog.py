from tkinter import *
from datetime import datetime
import pytz
import tkinter.messagebox
import csv
import os
from tzlocal import get_localzone

# Application to log and calculate minutes worked, both Straight Typing (ST) and
# Voice Recognition (VR) time


# Written by Erik Golke, 02/18/2017

# Get current date, format it to local time zone
date_format='%m/%d/%Y'
day_format='%m/%d/%Y/%I:%M:%S'
date = datetime.now(tz=pytz.utc)
local_tz = get_localzone()
date = date.astimezone(local_tz)


class MinuteLog:


    def __init__(self, master):
        """ Creates initial window of program along with entry fields and buttons"""

        self.root = master
        self.root.title('MinuteLog')
        self.root.geometry('350x350')

        self.regdict = Label(self.root, text="ST Dictation (Minutes: xxx.xx)", font=44)
        self. vrdict = Label(self.root, text="VR Dictation (Minutes: xxx.xx)", font=44)

        self.entry_regdict = Entry(self.root)
        self.entry_regdict.bind("<Return>", lambda x: self.log_minutes("Regular_Temp.csv", self.entry_regdict.get()))
        self.entry_vrdict = Entry(self.root)
        self.entry_vrdict.bind("<Return>", lambda x: self.log_minutes("VR_Temp.csv", self.entry_vrdict.get()))

        self.regdict.pack()
        self.entry_regdict.pack()

        self.button_1 = Button(self.root, text="Submit", font=44,
                               command=lambda: self.log_minutes("Regular_Temp.csv", self.entry_regdict.get()))
        self.button_1.pack(pady=(0, 20))

        self.vrdict.pack()
        self.entry_vrdict.pack()

        self.button_2 = Button(self.root, text="Submit", font=44,
                               command=lambda: self.log_minutes("VR_Temp.csv", self.entry_vrdict.get()))
        self.button_2.pack(pady=(0, 20))

        self.button_3 = Button(self.root, text="End of day - Log minutes",
                               font=44, command=self.save_temp_minutes)
        self.button_3.pack(pady=(10,0))

        self.button_4 = Button(self.root, text="ST Minute Summary",
                               font=44, command=lambda: self.summary("Regular_Dictation.csv"))
        self.button_4.pack(pady=(10, 0))

        self.button_5 = Button(self.root, text="VR Minute Summary",
                               font=44, command=lambda: self.summary("VR_Dictation.csv"))
        self.button_5.pack(pady=(10, 0))

        self.toolbar()
        self.status_bar()


    def summary(self, filepath):
        """ Creates summary window with entry fields and buttons"""

        self.minutesummary = Toplevel()
        self.minutesummary.title("Minute History")
        self.minutesummary.geometry('300x300')
        self.startdate = Label(self.minutesummary, text="Start Date (Format: mm/dd/yyyy)", font=44)
        self.enddate = Label(self.minutesummary, text="End Date(Format: mm/dd/yyyy)", font=44)
        self.startdate.pack()
        self.start_entry = Entry(self.minutesummary)
        self.start_entry.pack(pady=(0,20))
        self.end_entry = Entry(self.minutesummary)
        self.enddate.pack()
        self.end_entry.pack()
        self.button_entry = Button(self.minutesummary, text="Submit", font=44,
                              command=lambda: self.hour_adder(filepath, self.start_entry.get(), self.end_entry.get()))
        self.end_entry.bind("<Return>",
                            lambda x: self.hour_adder(filepath, self.start_entry.get(), self.end_entry.get()))
        self.start_entry.bind("<Return>",
                            lambda x: self.hour_adder(filepath, self.start_entry.get(), self.end_entry.get()))
        self.button_entry.pack()

    def status_bar(self):
        """ Creates window status bar"""
        status = Label(root, text="MinuteLog Version 2.3, Golke and Sons Solutions, Inc.", bd=1, relief=SUNKEN, anchor=W)
        status.pack(side=BOTTOM, fill=X)

    def toolbar(self):
        #os.system("start " + filename)
        """ Creates window toolbar"""
        self.menu = Menu(root)
        self.root.config(menu=self.menu)
        # File menu.
        self.subMenu = Menu(self.menu, tearoff=False)
        self.subMenu.add_command(label="View today's ST history",
                                 command=lambda: self.temp_adder("Regular_Temp.csv"))
        self.subMenu.add_command(label="View today's VR history",
                                 command=lambda: self.temp_adder("VR_Temp.csv"))
        self.subMenu.add_separator()
        self.menu.add_cascade(label="File", menu=self.subMenu)
        self.subMenu.add_command(label="Clear ST Summary History",
                                 command=lambda:self.history_clear("Regular_Dictation.csv"))
        self.subMenu.add_command(label="Clear VR Summary History",
                                 command=lambda:self.history_clear("VR_Dictation.csv"))
        self.subMenu.add_separator()
        self.subMenu.add_command(label="Exit", command=self.quit)

        # Edit menu.
        self.edit_menu = Menu(self.menu, tearoff=False)
        self.menu.add_cascade(label="Edit", menu=self.edit_menu)
        self.edit_menu.add_command(label="Edit today's ST history",
                                 command=lambda: self.open_file("Regular_Temp.csv"))
        self.edit_menu.add_command(label="Edit today's VR history",
                                 command=lambda: self.open_file("VR_Temp.csv"))
        self.edit_menu.add_separator()
        self.edit_menu.add_command(label="Edit ST Summary History",
                                 command=lambda: self.open_file("Regular_Dictation.csv"))
        self.edit_menu.add_command(label="Edit VR Summary History",
                                 command=lambda: self.open_file("VR_Dictation.csv"))

    def history_clear(self, filepath):
        """ Clears history, overwrites stated CSV file"""
        if filepath == "Regular_Dictation.csv":
            answer = tkinter.messagebox.askquestion("Clear ST history",
                                                    "Are you sure you want to clear all ST history?\n"
                                                    " (this action cannot be undone.)")
        elif filepath == "VR_Dictation.csv":
            answer = tkinter.messagebox.askquestion("Clear VR history",
                                                    "Are you sure you want to clear all VR history?\n"
                                                    " (this action cannot be undone.)")
        if answer == "yes":
            try:
                open(filepath, 'w').close()

            except IOError:
                tkinter.messagebox.showinfo("Error", "Error encountered while opening file.")


    def save_temp_minutes(self):
        """ Permanently saves the minutes stored in temp CSV files to main CSV files"""
        try:
            # Attempts to open files, initialize st_result counter to 0.
            with open("Regular_Temp.csv", "r", newline='') as infile:
                with open("Regular_Dictation.csv", "a", newline='') as outfile:
                    st_result = 0
                    reader = csv.reader(infile, delimiter=',')

                    for row in reader:
                        # Adds seconds in row to result.
                        st_result += int(row[1])

                    # Converts ST dictation time into minutes and seconds for display purposes.
                    st_min_result = int(st_result / 60)
                    st_sec_result = round(st_result % 60)

                    # Saves current date and result in a list, and writes that list to specified csv
                    st_data = [date.strftime(date_format), st_result]
                    # If sum of temp file is 0, don't save it.
                    if st_data[1] != 0:
                        writer = csv.writer(outfile, delimiter=',')
                        writer.writerow(st_data)

            infile.close()
            outfile.close()

            # Attempts to open files, initialize vr_result counter  to 0.
            with open("VR_Temp.csv", "r", newline='') as infile:
                with open("VR_Dictation.csv", "a", newline='') as outfile:
                    vr_result = 0
                    reader = csv.reader(infile, delimiter=',')

                    for row in reader:
                        # Add seconds in row to result
                        vr_result += int(row[1])

                    # Converts VR dictation time into minutes and seconds for display purposes
                    vr_min_result = int(vr_result / 60)
                    vr_sec_result = round(vr_result % 60)


                    # Saves current date and result in a list, and writes that list to specified CSV.
                    vr_data = [date.strftime(date_format), vr_result]
                    # If sum of temp file is 0, don't save it.
                    if vr_data[1] != 0:
                        writer = csv.writer(outfile, delimiter=',')
                        writer.writerow(vr_data)

            infile.close()
            outfile.close()
            # If no new data was entered, dont display message
            if st_data[1] or vr_data[1] != 0:
                # Displays total converted minutes and seconds in message box.
                tkinter.messagebox.showinfo(date.strftime(date_format),
                                            "Successfully saved " + str(st_min_result) + ":" + str(st_sec_result).zfill(2)
                                            + " ST minutes" + " and " + str(vr_min_result) + ":" + str(vr_sec_result).zfill(2)
                                            + " VR minutes for " + date.strftime(date_format))

            # Clears temp CSV files for future use.
            open("VR_Temp.csv", 'w').close()
            open("Regular_Temp.csv", 'w').close()


        except IOError:
            tkinter.messagebox.showinfo("Error", "Error encountered while opening file.")


    def log_minutes(self, filepath, entry_input):
        """ Logs input to a temporary csv file to be tallied at the end of day"""

        # Saves current time and input in list.
        data = [date.strftime(day_format), entry_input]
        try:
            # Ensures input is valid float.
            float(entry_input)
            try:
                with open(filepath, "a", newline='') as file:

                    # Converts inputted minute/second float  into seconds.
                    new_minutes = int(float(entry_input)) * 60
                    new_seconds = round(((float(entry_input) % 1) * 100))

                    # Replace entry_input with input converted to seconds.
                    data[1] = (new_minutes + new_seconds)

                    # Writes current time and converted input to CSV in seconds.
                    writer = csv.writer(file, delimiter=',')
                    writer.writerow(data)

                    # Converts total seconds into minutes and seconds for display.
                    min_result = int(data[1] / 60)
                    sec_result = round(data[1] % 60)

                    if filepath == "Regular_Temp.csv":

                        # Displays logged minutes and seconds in message box.
                        tkinter.messagebox.showinfo(date.strftime(day_format),
                                                    "Successfully logged " + str(min_result) + ":" +
                                                    str(sec_result).zfill(2) + " ST minutes")
                        file.close()
                        # Displays message box with current minute total
                        self.temp_adder("Regular_Temp.csv")
                        self.entry_regdict.delete(0, 'end')

                    elif filepath == "VR_Temp.csv":

                        # Displays logged minutes and seconds in message box
                        tkinter.messagebox.showinfo(date.strftime(day_format),
                                                    "Successfully logged " + str(min_result) + ":" +
                                                    str(sec_result).zfill(2) + " VR minutes")
                        file.close()
                        # Displays message box with current minute total
                        self.temp_adder("VR_Temp.csv")
                        self.entry_vrdict.delete(0, 'end')
            except IOError:
                tkinter.messagebox.showinfo("Error", "Error encountered while opening file.")
        except ValueError:
            tkinter.messagebox.showinfo("Error", "Input must be numerical")

    def temp_adder(self, filepath):
        """ Adds up all hours in temporary csv file"""

        # Attempts to open specified file.
        try:
            with open(filepath, "r", newline='') as file:
                result = 0
                reader = csv.reader(file, delimiter=',')
                minutes = []

                for row in reader:
                    # If a blank line is encountered, skip it.
                    try:
                        # Adds seconds in row to result, and appends date and time in a dict to minutes list.
                        result += int(row[1])
                        # Converts seconds stored in CSV to minutes and seconds for cleaner output, and adds to dict
                        converted_string = self.seconds_to_minutes(row[1])
                        final = ':'.join(map(str, converted_string))
                        minutes.append({"Date": row[0], " Minutes": final})
                    except ValueError:
                        continue

                # Converts elements in list to strings and appends to one string separated by newline
                Minutes = '\n'.join(map(str, minutes))

                # Strips unwanted punctuation from string for palatable output.
                Minutes = self.stripper(Minutes)

                # Converts seconds to minutes and seconds for output.
                min_result = int(result / 60)
                sec_result = round(result % 60)

                # Displays minutes and seconds in a message box.
                if filepath == "Regular_Temp.csv":
                    tkinter.messagebox.showinfo("Minute Summary", Minutes + "\nTotal ST Minutes for today: " + str(min_result) + ":" + str(sec_result).zfill(2))
                else:
                    tkinter.messagebox.showinfo("Minute Summary",
                                                Minutes + "\nTotal VR Minutes for today: " + str(min_result) + ":" + str(sec_result).zfill(2))
                file.close()
        except IOError:
            tkinter.messagebox.showinfo("Error", "Error encountered while opening file.")

    def hour_adder(self, filepath, start, end):
        """ Takes start date and end date, calculates number of minutes in between"""

        # Attempts to open specified file.
        try:
            with open(filepath, "r", newline='') as file:
                result = 0
                reader = csv.reader(file, delimiter=',')
                minutes = []
                flag = False

                # If only one date given, or if both dates match, only display results of that date.
                if len(start) == 0 or len(end) == 0 or start == end:
                    for row in reader:
                        # If a blank line is encountered, skip it.
                        try:
                            if row[0] == end or row[0] == start:
                                result += float(row[1])
                                # Converts seconds stored in CSV to minutes and seconds for cleaner output, and adds to dict
                                converted_string = self.seconds_to_minutes(row[1])
                                final = ':'.join(map(str, converted_string))

                                minutes.append({"Date": row[0], " Minutes": final})
                        except ValueError:
                            continue
                # If valid date range given, add seconds between the two dates.
                else:
                    for row in reader:
                        # If a blank line is encountered, skip it.
                        try:
                            if row[0] == start:
                                result += float(row[1])
                                # Converts seconds stored in CSV to minutes and seconds for cleaner output, and adds to dict
                                converted_string = self.seconds_to_minutes(row[1])
                                final = ':'.join(map(str, converted_string))
                                minutes.append({"Date": row[0], " Seconds": final})
                                flag = True
                                continue

                            if flag:
                                result += float(row[1])
                                # Converts seconds stored in CSV to minutes and seconds for cleaner output, and adds to dict
                                converted_string = self.seconds_to_minutes(row[1])
                                final = ':'.join(map(str, converted_string))
                                minutes.append({"Date": row[0], " Seconds": final})
                            if row[0] == end:
                                flag = False
                                break
                        except ValueError:
                            continue
                file.close()
        except IOError:
            tkinter.messagebox.showinfo("Error", "Error encountered while opening file.")

        # Converts elements in list to strings and appends to one string separated by newline
        Minutes = '\n'.join(map(str, minutes))

        # Strips unwanted punctuation from string for palatable output.
        Minutes = self.stripper(Minutes)

        # Clears entry boxes of any text for aesthetic reasons.
        self.start_entry.delete(0, 'end')
        self.end_entry.delete(0, 'end')

        # Convert seconds to minutes and seconds for output
        min_result = int(result / 60)
        sec_result = round(result % 60)

        # Displays total minutes and seconds in message box.
        if filepath == "Regular_Dictation.csv":
            tkinter.messagebox.showinfo("Minute Summary", Minutes + "\nTotal ST Minutes: " + str(min_result) + ":"
                                        + str(sec_result).zfill(2))
        else:
            tkinter.messagebox.showinfo("Minute Summary", Minutes + "\nTotal VR Minutes: " + str(min_result) + ":"
                                        + str(sec_result).zfill(2))


    def stripper(self, string):
        """ Strips unwanted punctuation from string"""
        for c in ['{', '}', '\'', '\'']:
            if c in string:
                string = string.replace(c, '')
        return string

    def seconds_to_minutes(self, row):
        """ Converts seconds stored in CSV to minutes for clean output"""
        converted_minutes = int(float(row) / 60)
        converted_seconds = round(float(row) % 60)
        converted_string = (converted_minutes, str(converted_seconds).zfill(2))
        return converted_string

    def quit(self):
        """ Quits the GUI, closing any open windows, and saves temp data to file"""
        self.save_temp_minutes()
        self.root.destroy()

    def open_file(self, filepath):
        """ Opens specified file with default program"""
        tkinter.messagebox.showinfo("Reminder", "Please note: this program saves minutes as seconds.")
        os.startfile(filepath)

# Sets the whole dirty ballgame in motion.
root = Tk()
app = MinuteLog(root)
root.protocol("WM_DELETE_WINDOW", app.quit)
root.mainloop()
