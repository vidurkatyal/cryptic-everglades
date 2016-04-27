import shutil
import os
import threading
import cairo
import math
import time

import gi
gi.require_version('Poppler', '0.18')
from gi.repository import Poppler
from gi.repository import Gtk
from gi.repository import GObject


class PDF_Doc:
    """Class for handling PDF documents."""

    def __init__(self, _file, numfiles, tmp_dir):

        self._file = os.path.abspath(_file)
        (self.path, self.filename) = os.path.split(self._file)
        (self.filename, self.ext) = os.path.splitext(self.filename)
        file_prefix = 'file://'

        self.filenum = numfiles + 1
        self.mtime = os.path.getmtime(_file)
        self.copyname = os.path.join(tmp_dir, '%02d_' % self.filenum +
                                                  self.filename + '.pdf')
        shutil.copy(self._file, self.copyname)
        self.document = Poppler.Document.new_from_file(file_prefix + self.copyname, None)
        self.numpages = self.document.get_n_pages()

class PDF_Renderer(threading.Thread, GObject.GObject):
    "Class for rendering thumbnails of pages."

    def __init__(self, model, pdfqueue):
        threading.Thread.__init__(self)
        GObject.GObject.__init__(self)
        self.model = model
        self.pdfqueue = pdfqueue
        self.quit = False

    def run(self):
        for idx, row in enumerate(self.model):
            if self.quit:
                return
            if not row[1]:
                try:
                    nfile = row[2]
                    npage = row[3]
                    pdfdoc = self.pdfqueue[nfile - 1]
                    page = pdfdoc.document.get_page(npage-1)
                    w, h = page.get_size()
                    thumbnail = cairo.ImageSurface(cairo.FORMAT_ARGB32,
                                                   int(w/2.0),
                                                   int(h/2.0))
                    cr = cairo.Context(thumbnail)
                    cr.scale(1.0/2.0, 1.0/2.0)
                    page.render(cr)
                    time.sleep(0.003)
                    GObject.idle_add(self.emit,'update_thumbnail',
                                     idx, thumbnail, priority=GObject.PRIORITY_LOW)
                except Exception as e:
                    print(e)



class CellRendererImage(Gtk.CellRenderer):
    __gproperties__ = {
            "image": (GObject.TYPE_PYOBJECT, "Image", "Image",
                      GObject.PARAM_READWRITE),
            "width": (GObject.TYPE_FLOAT, "Width", "Width",
                      0., 1.e4, 0., GObject.PARAM_READWRITE),
            "height": (GObject.TYPE_FLOAT, "Height", "Height",
                       0., 1.e4, 0., GObject.PARAM_READWRITE),
            "rotation": (GObject.TYPE_INT, "Rotation", "Rotation",
                         0, 360, 0, GObject.PARAM_READWRITE),
            "scale": (GObject.TYPE_FLOAT, "Scale", "Scale",
                      0.01, 100., 1., GObject.PARAM_READWRITE),
    }

    def __init__(self):
        Gtk.CellRenderer.__init__(self)
        self.th1 = 2. # border thickness
        self.th2 = 3. # shadow thickness

    def get_geometry(self):

        rotation = int(self.rotation) % 360
        rotation = round(rotation / 90) * 90
        if not self.image:
            w0 = w1 = self.width / 2.0
            h0 = h1 = self.height / 2.0
        else:
            w0 = self.image.get_width()
            h0 = self.image.get_height()
            if rotation == 90 or rotation == 270:
                w1, h1 = h0, w0
            else:
                w1, h1 = w0, h0

        scale = 2.0 * self.scale
        w2 = int(scale * w1)
        h2 = int(scale * h1)
        
        return w0,h0,w1,h1,w2,h2,rotation

    def do_set_property(self, pspec, value):
        setattr(self, pspec.name, value)

    def do_get_property(self, pspec):
        return getattr(self, pspec.name)

    def do_render(self, window, widget, background_area, cell_area, expose_area):
        if not self.image:
            return

        w0,h0,w1,h1,w2,h2,rotation = self.get_geometry()
        th = int(2*self.th1+self.th2)
        w = w2 + th
        h = h2 + th

        x = cell_area.x
        y = cell_area.y
        if cell_area and w > 0 and h > 0:
            x += self.get_property('xalign') * \
                 (cell_area.width - w - self.get_property('xpad'))
            y += self.get_property('yalign') * \
                 (cell_area.height - h - self.get_property('ypad'))

        window.translate(x,y)

        x = 0
        y = 0

        #shadow
        window.set_source_rgb(0.5, 0.5, 0.5)
        window.rectangle(th, th, w2, h2)
        window.fill()

        #border
        window.set_source_rgb(0, 0, 0)
        window.rectangle(0, 0, w2+2*self.th1, h2+2*self.th1)
        window.fill()

        #image
        window.set_source_rgb(1, 1, 1)
        window.rectangle(self.th1, self.th1, w2, h2)
        window.fill_preserve()
        window.clip()

        window.translate(self.th1,self.th1)
        scale = 2.0 * self.scale
        window.scale(scale, scale)
        window.translate(-x,-y)
        if rotation > 0:
            window.translate(w1/2,h1/2)
            window.rotate(rotation * math.pi / 180)
            window.translate(-w0/2,-h0/2)

        window.set_source_surface(self.image)
        window.paint()

    def do_get_size(self, widget, cell_area=None):
        x = y = 0
        w0,h0,w1,h1,w2,h2,rotation = self.get_geometry()
        th = int(2*self.th1+self.th2)
        w = w2 + th
        h = h2 + th

        if cell_area and w > 0 and h > 0:
            x = self.get_property('xalign') * \
                (cell_area.width - w - self.get_property('xpad'))
            y = self.get_property('yalign') * \
                (cell_area.height - h - self.get_property('ypad'))
        w += 2 * self.get_property('xpad')
        h += 2 * self.get_property('ypad')
        return int(x), int(y), w, h