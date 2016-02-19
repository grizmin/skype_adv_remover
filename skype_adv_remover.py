# coding:utf-8

import _winreg as wreg
import os
import re
import tkMessageBox
from Tkinter import *

import psutil

""" Quick and dirty skype advertisements remover """


class Build_GUI(object):
    """ Build UI """

    def __init__(self, root='root'):
        self.root = root
        root.title('Skype adv remover by Your Mom')
        self.root.resizable(0, 0)
        self.restartvar = IntVar()

        self.frame_usernames = Frame(root)
        self.frame_button = Frame(root)
        self.B = Button(self.frame_button, text="OK", command=self.ok_button, width=7, height=3)
        self.L = Label(self.root, text="select skype profile to patch", width=26, relief=RIDGE, padx=18, pady=3)
        self.R = Checkbutton(self.frame_button, text="Restart Skype", variable=self.restartvar, onvalue=1, offvalue=0)
        self.R.select()

        self.frame_usernames.grid(row=1, column=0)
        self.frame_button.grid(row=1, column=1)

        self.L.grid(row=0, column=0, sticky=NW, columnspan=2)
        self.B.grid(row=0, column=0, pady=5, padx=15, sticky=NSEW)
        self.R.grid(row=1, column=0, padx=0, pady=0)

        
    def center_window(self, all_skype_profiles, width=223, height=0):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        height += (len(all_skype_profiles) * 20)
        if height < 120: height = 120
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        self.root.geometry('%dx%d+%d+%d' % (width, height, x, y))

    def draw_profile_checkbuttons(self, skype_profiles):
        self.genvars = {}
        rows = 1
        for i in skype_profiles:
            self.genvars[i] = IntVar(name=i)
            i = Checkbutton(self.frame_usernames, text="%s" % i[:15], variable=i, onvalue=1, offvalue=0)
            i.grid(row=rows, column=0, sticky=NW)
            i.select()
            rows += 1

    def ok_button(self):
        pass


class Action(Build_GUI):
    def __init__(self, root):
        Build_GUI.__init__(self, root)
        all_skype_profiles = self.find_all_skype_profiles_winreg()
        if not all_skype_profiles:
            all_skype_profiles = self.find_all_skype_profiles()
        self.center_window(all_skype_profiles)
        self.draw_profile_checkbuttons(all_skype_profiles)        

    def ok_button(self):
        """ executes all changes """
        skype_profile_to_patch = []
        for i in self.genvars.keys():
            if self.genvars[i].get():
                skype_profile_to_patch.append(i)
        if not len(skype_profile_to_patch):
            tkMessageBox.showinfo("Error", "No skype profile was selected!")

        else:
            for skype_profile in skype_profile_to_patch:
                self.change_adv_config(skype_profile)
                self.add_restricted_site()
            if self.restartvar.get():
                self.restart_skype()
            tkMessageBox.showinfo("Congratulations!", "Your skype will display no more annoying advertisements.")

    def get_pid(self, pname):
        """ takes string process name, returns integer process id"""
        process = filter(lambda p: p.name() == '%s' % pname, psutil.process_iter())
        for i in process:
            return i.pid

    def find_all_skype_profiles(self):
        """ lists all profile names """
        skype_profiles = []
        dirwalk = os.walk(os.sep.join([os.environ['APPDATA'], 'Skype']))
        for root, dirs, files in dirwalk:
            for filename in files:
                if 'config.xml' in filename:
                    skype_profiles.append(os.path.split(root)[1])
        return skype_profiles

    def find_all_skype_profiles_winreg(self):
        """ lists all profile names using winreg """
        skype_profiles = []
        try:
            skype_regentry = wreg.OpenKey(wreg.HKEY_CURRENT_USER, "Software\\Skype\\Phone\\Users")
            for i in range(20):
                try:
                    skype_profile_name = wreg.EnumKey(skype_regentry, i)
                    if skype_profile_name:
                        skype_profiles.append(skype_profile_name)
                except WindowsError as werror:
                    pass
        except WindowsError:
            return 0            
        finally:
                wreg.CloseKey(skype_regentry)            
        return skype_profiles

    def change_adv_config(self, skype_profile_name):
        """ removes advertisements from config.xml """
        skype_config_path = os.sep.join([os.environ['APPDATA'], 'Skype', skype_profile_name, 'config.xml'])
        adv_match = '(?<=\<AdvertPlaceholder\>)\d(?=\<\/AdvertPlaceholder\>)'
        match = re.compile(adv_match)

        try:
            with open(skype_config_path, mode='r+') as fs:
                skype_config = fs.read()
                skype_config = match.sub('0', skype_config)
                fs.seek(0)
                fs.write(skype_config)

        except Exception as e:
            tkMessageBox.showinfo("Error", "Your Skype profile name \'%s\' was not modified." % skype_config_path)
            raise
        print skype_config_path, "was changed."
        return

    def add_restricted_site(self):
        """ adds apps.skype.com in restricted sites """
        skype_domain = 'skype.com'
        skype_apps_domain = 'apps'
        key = wreg.CreateKey(wreg.HKEY_CURRENT_USER,
                             "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Internet Settings\\ZoneMap\\Domains\\{0}".format(
                                 skype_domain, skype_apps_domain))
        wreg.CloseKey(key)
        key = wreg.CreateKey(wreg.HKEY_CURRENT_USER,
                             "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Internet Settings\\ZoneMap\\Domains\\{0}\\{1}".format(
                                 skype_domain, skype_apps_domain))
        wreg.SetValueEx(key, '*', 0, wreg.REG_DWORD, 4)
        wreg.CloseKey(key)

    def restart_skype(self):
        """ restart skype """
        skype_regentry = wreg.OpenKey(wreg.HKEY_CURRENT_USER, "Software\\Skype\\Phone")
        skype_executable_path = wreg.QueryValueEx(skype_regentry, "SkypePath")[0]
        if not skype_executable_path: tkMessageBox.showinfo("Error", "Your Skype installation was not found")
        skype_process = self.get_pid("Skype.exe")
        if skype_process:
            skype_executable = psutil.Process(skype_process)
            skype_executable.terminate()
            if os.path.exists(skype_executable_path):
                os.startfile(skype_executable_path)
        else:
            print("Skype is not running. Starting it.")
            if os.path.exists(skype_executable_path):
                s.startfile(skype_executable_path)


if __name__ == '__main__':
    root = Tk()
    skype_adv_remover = Action(root)
    root.mainloop()
