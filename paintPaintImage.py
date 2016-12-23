import ctypes
import tkMessageBox
import tkFileDialog
import sys
import multiprocessing
import pyHook
import pythoncom 
from Tkinter import *
from PIL import ImageTk, Image
from collections import OrderedDict
from time import sleep

__author__ = "Toni Vuoristo"
__version__ = '1.0.1'
__email__ = "greatshader@hotmail.com"


class ReadAbortKey(object):
    """
        A class interrupter to watch for abort key via hooking to keyboard

        Spawns as second process via multiprocessing 

        FIX: Zombie process problem
    """
    def __init__(self, queue):
        self.queue = queue
        self.abort_key = 'a'

    def OnKeyboardEvent(self, event):
        """
            Actual event hooked to HookManager keyboard

            reads and handles the signaling of the kill signal for main execution
        """
        c = chr(event.Ascii)
        if c == self.abort_key:
            self.queue.put(self.abort_key)
        return True

    def loopReadKeyBoard(self):
        """
            Hijack the keyboard to read everything typed on it

            on a lookout for abort key
        """
        hooks_manager = pyHook.HookManager()
        hooks_manager.KeyDown = self.OnKeyboardEvent
        hooks_manager.HookKeyboard()
        pythoncom.PumpMessages()   
        

class Main(object):
    """
        Main class that holds all main functions handling the drawing on paint

    """
    def __init__(self, queue):
        self.queue = queue
        self.base = Tk()
        self.base.title("Paint image")
        self.base.resizable(0, 0)
        self.base.config(padx=60, pady=60)
        self.load_img = Button(self.base, text='Setup image',
                               command=self.imagePreview)
        self.load_img.config(padx=10, pady=10)
        self.load_img.grid(sticky=W+E+S+N)
        self.begin = Button(self.base, text="Begin painting",
                            command=self.beginDrawing)
        self.begin.config(padx=10, pady=10)
        
        self.begin.grid(sticky=W+E+S+N)
        self.start_loc = (300, 300)
        self.image_path = ''
        self.image = None
        self.img_scan = None
        self.prog_var = StringVar()
        self.prog_var.set('0 / 0')
        self.prog_dis = Label(self.base, textvariable=self.prog_var)
        self.prog_dis.grid()

    def load_image(self):
        """
            Handles the image loading and making sure that it is indeed an image

        """
        filename = tkFileDialog.askopenfilename()
        if filename == '':
            self.top.lift()
            return None
        try:
            # Keep the original image intact incase needed later on
            self.image = ImageTk.PhotoImage(Image.open(filename))
            # New variable to hold the converted image
            self.img_scan = Image.open(filename).convert('RGB')
            self.img_label.config(image=self.image)
            # Store the path incase needed later
            self.image_path = filename
        except IOError:
            tkMessageBox.showerror('ImageError', "Imagetype not valid!")

        self.top.lift()

    def abortPainting(self):
        """
            Function that is called upon wanting to abort the execution

            abort signal comes from other process via queue 
        """
        try:
            if self.queue.get(block=False) == 'a':
                tkMessageBox.showwarning('Abort', "Painting aborted!") 
                return True
        except Exception:
            return False

    def beginDrawing(self):
        """
            Main function of drawing

            handles chopping up the image to each individual pixels
            and append every same rgb pixel under same roof
        
        """
        if self.image is None:
            tkMessageBox.showerror('ImageError',
                                   "Start by loading an image!")
            return None
        sleep(3)
        final = OrderedDict()
        size_x, size_y = self.img_scan.size
        for y in xrange(size_y):
            for x in xrange(size_x):
                rgb = self.img_scan.getpixel((x, y))
                key = '.'.join([str(t) for t in rgb])
                if key in final:
                    final[key].append((x, y))
                else:
                    final[key] = [] 
                    final[key].append((x, y))

        length = len(final)
        cnt = 0
        for key, value in final.iteritems():
            self.prog_var.set('{0} / {1}'.format(cnt, length))
            self.prog_dis.update()
            if self.abortPainting():
                return None
            # Set the next color in loop
            self.goToModColors(key.split('.'))
            # Start painting with given color
            sx, sy = self.start_loc 
            for paint in value:
                if self.abortPainting():
                    return None
                # Paint, paint and more paint!
                x, y = paint 
                ctypes.windll.user32.SetCursorPos(sx+x, sy+y)
                sleep(0.002)
                ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
                ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)
            cnt += 1

            
            
    def goToModColors(self, rgb=['255', '255', '255']):
        """
            A function when call'd, will locate it self to 'edit colors' tab
            to change to given color

            contains mostly just position coordinates and input from keyboard and mouse
        """
        # Press the Edit colors
        ctypes.windll.user32.SetCursorPos(1047, 80)
        sleep(0.2)
        ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
        ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)
        # Go through r, g, b to apply the given colors
        for s, p in zip((593, 615, 638), rgb):
            ctypes.windll.user32.SetCursorPos(1165, s)
            sleep(0.1)
            ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
            ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)
            for d in xrange(3):
                # Press delete 3 times to get rid of default stuff
                ctypes.windll.user32.keybd_event(int('8', 16), 0, 2, 0)
                ctypes.windll.user32.keybd_event(int('8', 16), 0, 0, 0)
            #Apply the color
            for c in p:
                ctypes.windll.user32.keybd_event(ord(c), 0, 2, 0)
                ctypes.windll.user32.keybd_event(ord(c), 0, 0, 0)
        # Press the OK
        ctypes.windll.user32.SetCursorPos(772, 663)
        sleep(0.1)
        ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
        ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)
        sleep(0.1)

    def imagePreview(self):
        """
            Creates a window, to allow upload of the about to be painted image
            and preview it.
        
        """
        if hasattr(self, 'top'):
            self.top.destroy()
        self.top = Toplevel()
        self.frame_p = LabelFrame(self.top, text='Load/Preview image')
        self.frame_p.config(padx=60, pady=60)
        self.frame_p.grid(padx=5, pady=5)
        load_img = Button(self.frame_p, text='Load image...',
                          command=self.load_image)
        load_img.grid(padx=20, pady=5, sticky=W+E+N)
        self.img_label = Label(self.frame_p, image=self.image \
                               if self.image is not None else None)
        self.img_label.grid()
        

    def mainloop(self):
        "Self explanatory"
        self.base.mainloop()



if __name__ == '__main__':
    q = multiprocessing.Queue()
    abort = ReadAbortKey(q)
    abort_process = multiprocessing.Process(target=abort.loopReadKeyBoard)
    abort_process.daemon = True
    abort_process.start()
    main = Main(q)
    main.mainloop()
