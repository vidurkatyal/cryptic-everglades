try:
    import gi
    gi.require_version('Gtk', '3.0')
    from gi.repository import Gtk
except ValueError, e:
    raise e

import sys
import mimetypes
import tempfile
import copy
from gi.repository import Gdk
from PyPDF2 import PdfFileWriter, PdfFileReader

from helpers import *



class PDFShuffler:
    """Class for PDFShuffler."""
    MODEL_ROW_INTERN = 1001
    MODEL_ROW_EXTERN = 1002
    TEXT_URI_LIST = 1003
    MODEL_ROW_MOTION = 1004
    TARGETS_ICONVIEW = [Gtk.TargetEntry.new('MODEL_ROW_INTERN', Gtk.TargetFlags.SAME_WIDGET, MODEL_ROW_INTERN),
                  Gtk.TargetEntry.new('MODEL_ROW_EXTERN', Gtk.TargetFlags.OTHER_APP, MODEL_ROW_EXTERN),
                  Gtk.TargetEntry.new('MODEL_ROW_MOTION', 0, MODEL_ROW_MOTION)]
    TARGETS_SW = [Gtk.TargetEntry.new('text/uri-list', 0, TEXT_URI_LIST),
                  Gtk.TargetEntry.new('MODEL_ROW_EXTERN', Gtk.TargetFlags.OTHER_APP, MODEL_ROW_EXTERN)]


    def __init__(self):
        # Make a temporary working directory.
        self.tmp_dir = tempfile.mkdtemp("shuffler_working_dir")
        os.chmod(self.tmp_dir, 0o700)


        # Build UI from glade file and connect handlers.
        self.uiBuilder = Gtk.Builder()
        self.uiBuilder.add_from_file("GUI.glade")
        self.uiBuilder.connect_signals(self)


        # Create the main window.
        self.window = self.uiBuilder.get_object("main_window")
        self.window.set_title("PDFShuffler")
        self.window.set_border_width(0)
        self.window.set_default_size(
            min(700, Gdk.Screen.get_default().get_width() / 2),
            min(600, Gdk.Screen.get_default().get_height() - 50)
            )
        self.window.connect('delete-event', self.close_window)


        # Settings for ScrolledWindow
        self.sw = self.uiBuilder.get_object('scrolledwindow')
        self.sw.drag_dest_set(Gtk.DestDefaults.MOTION |
                              Gtk.DestDefaults.HIGHLIGHT |
                              Gtk.DestDefaults.DROP |
                              Gtk.DestDefaults.MOTION,
                              self.TARGETS_SW,
                              Gdk.DragAction.COPY |
                              Gdk.DragAction.MOVE)

        align = Gtk.Alignment.new(0.5, 0.5, 0, 0)
        self.sw.add_with_viewport(align)

        # Create the progress bar
        self.progress_bar = self.uiBuilder.get_object('progressbar')
        self.progress_bar_timeout_id = 0

        # Model to hold individual pages of a pdf.
        # (In a tabular format i.e. rows and columns)
        # Defining columns
        self.model = Gtk.ListStore(str,         # 0. Thumbnail caption
                                   GObject.TYPE_PYOBJECT, # 1.Cached page image
                                   int,         # 2. Document number
                                   int,         # 3. Page number
                                   float,       # 4. Scale
                                   str,         # 5. Document filename
                                   int,         # 6. Rotation angle
                                   float,       # 7. Crop left
                                   float,       # 8. Crop right
                                   float,       # 9. Crop top
                                   float,       # 10. Crop bottom
                                   float,       # 11. Page width
                                   float)       # 12. Page height

        self.zoom_set(-14)

        # Settings for thumbnails
        self.iconview_col_width = 300
        self.iconview = Gtk.IconView(self.model)
        self.iconview.clear()
        self.iconview.set_item_width(-1)

        self.cellthumb = CellRendererImage()
        self.iconview.pack_start(self.cellthumb, False)
        self.iconview.set_cell_data_func(self.cellthumb, self.set_cell_data, None)



        # Selection settings for iconview.
        self.iconview.set_text_column(0)
        self.iconview.set_selection_mode(Gtk.SelectionMode.MULTIPLE)
        self.iconview.enable_model_drag_source(Gdk.ModifierType.BUTTON1_MASK,
                                               self.TARGETS_ICONVIEW,
                                               Gdk.DragAction.COPY |
                                               Gdk.DragAction.MOVE)
        self.iconview.enable_model_drag_dest(self.TARGETS_ICONVIEW,
                                             Gdk.DragAction.DEFAULT)

        # TODO: Connect drag and drop events to handlers





        align.add(self.iconview)

        # Change iconview color background
        style_context_sw = self.sw.get_style_context()
        color_selected = self.iconview.get_style_context().get_background_color(Gtk.StateFlags.SELECTED)
        color_prelight = color_selected.copy()
        color_prelight.alpha = 0.3
        for state in (Gtk.StateFlags.NORMAL, Gtk.StateFlags.ACTIVE):
           self.iconview.override_background_color(state, style_context_sw.get_background_color(state))
        self.iconview.override_background_color(Gtk.StateFlags.SELECTED, color_selected)
        self.iconview.override_background_color(Gtk.StateFlags.PRELIGHT, color_prelight)


        # Connecting size change handler to window.
        self.window.connect('size_allocate', self.window_resize)
        self.window.show_all()
        self.progress_bar.hide()


        # Variables
        self.import_directory = self.export_directory = os.getenv('HOME')
        self.pdfqueue = []
        self.numfiles = 0
        self.iconview_auto_scroll_direction = 0
        self.iconview_auto_scroll_timer = None


        GObject.type_register(PDF_Renderer)
        GObject.signal_new('update_thumbnail', PDF_Renderer,
                           GObject.SignalFlags.RUN_FIRST, None,
                           [GObject.TYPE_INT, GObject.TYPE_PYOBJECT])
        self.rendering_thread = 0


    def close_window(self, widget, event=None, data=None):
        """Handler for closing the application."""

        if os.path.isdir(self.tmp_dir):
            shutil.rmtree(self.tmp_dir)

        if Gtk.main_level():
            Gtk.main_quit()
        else:
            sys.exit(0)

    def window_resize(self, window, event):
        """Handler for window resize."""

        den = 10 * (self.iconview_col_width + self.iconview.get_column_spacing() * 2)
        col_num = 9 * window.get_size()[0] / den
        self.iconview.set_columns(col_num)

    def import_pdf(self, widget, data=None):
        """Handler for importing new pdf(s)."""

        file_import = Gtk.FileChooserDialog(title="Choose PDF file(s) to import..",
                                            parent=self.window,
                                            action=Gtk.FileChooserAction.OPEN,
                                            buttons=(Gtk.STOCK_CANCEL,
                                                     Gtk.ResponseType.CANCEL,
                                                     Gtk.STOCK_OPEN,
                                                     Gtk.ResponseType.OK
                                                    )
                                            )

        file_import.set_select_multiple(True)
        file_import.set_current_folder(self.import_directory)
        filter_pdf = Gtk.FileFilter()
        filter_pdf.set_name('PDF files')
        filter_pdf.add_mime_type('application/pdf')
        file_import.add_filter(filter_pdf)

        response = file_import.run()
        if response == Gtk.ResponseType.OK:
            errors = ""
            for _file in file_import.get_filenames():
                (path, filename) = os.path.split(_file)
                self.import_directory = self.export_directory = path
                file_mimetype = mimetypes.guess_type(_file)[0]
                if file_mimetype != 'application/pdf':
                    errors += "File %s not supported! Not importing...\n" % filename
                    continue

                rv = self.add_pdf(_file)
                if not rv:
                    errors += "Error importing file %s...\n" % filename

            if errors:
                self.error_message_dialog(errors)

        file_import.destroy()

    def export_pdf(self, widget=None, data=None):
        """Handler for exporting the pdf."""

        # Checking whether the application currently has some pdfs
        if not self.pdfqueue:
            error = "Error! No file imported."
            self.error_message_dialog(error)
            return

        file_export = Gtk.FileChooserDialog(title="Export PDF...",
                                            parent=self.window,
                                            action=Gtk.FileChooserAction.SAVE,
                                            buttons=(Gtk.STOCK_CANCEL,
                                                     Gtk.ResponseType.CANCEL,
                                                     Gtk.STOCK_SAVE,
                                                     Gtk.ResponseType.OK
                                                    )
                                            )

        file_export.set_do_overwrite_confirmation(True)
        file_export.set_current_folder(self.export_directory)
        filter_pdf = Gtk.FileFilter()
        filter_pdf.set_name('PDF files')
        filter_pdf.add_mime_type('application/pdf')
        file_export.add_filter(filter_pdf)

        response = file_export.run()
        if response == Gtk.ResponseType.OK:
            _file = file_export.get_filename()
            (path, filename) = os.path.split(_file)
            (filename, ext) = os.path.splitext(filename)
            if ext.lower() != '.pdf':
                _file += '.pdf'

            try:
                self.export_to_file(_file)
            except Exception, error: # To catch exceptions like OSError
                self.error_message_dialog(error)

        file_export.destroy()


    def add_pdf(self, _file, startpage=None, endpage=None,
                            angle=0.0, crop=[0.0, 0.0, 0.0, 0.0]):
        """Function to add pdf to the application."""

        pdfdoc = None
        rv = False

        for pdf in self.pdfqueue:
            if os.path.isfile(pdf.filename):
                if os.path.samefile(_file, pdf.filename):
                    if os.path.getmtime(_file) is pdf.mtime:
                        pdfdoc = pdf
                        break


        if not pdfdoc:
            pdfdoc = PDF_Doc(_file, self.numfiles, self.tmp_dir)
            if not pdfdoc:
                return rv

            self.numfiles = pdfdoc.filenum
            self.pdfqueue.append(pdfdoc)


        start = 1
        end = pdfdoc.numpages
        if startpage:
            start = min(end, max(1, startpage))
        if endpage:
            end = max(start, min(end, endpage))


        for page_num in range(start, end + 1):
            caption = pdfdoc.filename + '\nPage ' + str(page_num)
            page = pdfdoc.document.get_page(page_num-1)
            width, height = page.get_size()
            iter = self.model.append((caption,
                                      None,
                                      pdfdoc.filenum,
                                      page_num,
                                      self.zoom_scale,
                                      pdfdoc._file,
                                      angle,
                                      crop[0],
                                      crop[1],
                                      crop[2],
                                      crop[3],
                                      width,
                                      height
                                    ))
            self.update_geometry(iter)
            rv = True

        self.reset_iconview_width()
        if rv:
            GObject.idle_add(self.render)
        return rv


    def export_to_file(self, _file):
        """Function to export the pdf to file."""

        pdf_output = PdfFileWriter()
        pdf_input = []

        for pdfdoc in self.pdfqueue:
            pdfdoc_file = PdfFileReader(open(pdfdoc.copyname, 'rb'))
            pdf_input.append(pdfdoc_file)

        for row in self.model:
            filenum, page_num = row[2], row[3]
            current_page = copy.copy(pdf_input[filenum-1].getPage(page_num-1))
            # TODO: Code for rotation and cropping



            pdf_output.addPage(current_page)


        pdf_output.write(open(_file, 'wb'))



    def error_message_dialog(self, msg):
        """Function to display error messages."""

        dialog = Gtk.MessageDialog(flags=Gtk.DialogFlags.MODAL,
                                          type=Gtk.MessageType.ERROR,
                                          parent=self.window,
                                          message_format=str(msg),
                                          buttons=Gtk.ButtonsType.OK)
        response = dialog.run()
        if response == Gtk.ResponseType.OK:
            dialog.destroy()


    def set_cell_data(self, column, cell, model, iter, data=None):
        """Function to set cell data."""

        cell.set_property('image', model.get_value(iter,1))
        cell.set_property('scale', model.get_value(iter,4))
        cell.set_property('rotation', model.get_value(iter,6))
        cell.set_property('cropL', model.get_value(iter,7))
        cell.set_property('cropR', model.get_value(iter,8))
        cell.set_property('cropT', model.get_value(iter,9))
        cell.set_property('cropB', model.get_value(iter,10))
        cell.set_property('width', model.get_value(iter,11))
        cell.set_property('height', model.get_value(iter,12))



    def reset_iconview_width(self, renderer=None):
        """Function to reconfigure the width of the iconview columns."""

        if not self.model.get_iter_first(): # Check if the model is empty
            return

        max_w = 10 + int( max(row[4]*row[11]*(1.0-row[7]-row[8]) for row in self.model) )
        if max_w != self.iconview_col_width:
            self.iconview_col_width = max_w
            self.iconview.set_item_width(-1)

            self.window_resize(self.window, None)

    def render(self):
        """Function to render the thumbnails."""

        if self.rendering_thread:
            self.rendering_thread.quit = True
            self.rendering_thread.join()

        self.rendering_thread = PDF_Renderer(self.model, self.pdfqueue)
        self.rendering_thread.connect('update_thumbnail', self.update_thumbnail)
        self.rendering_thread.start()

        if self.progress_bar_timeout_id:
            GObject.source_remove(self.progress_bar_timeout_id)
        self.progress_bar_timout_id = GObject.timeout_add(50, self.progress_bar_timeout)

    def update_geometry(self, iter):
        """Function to recompute the width and height of rotated pages."""

        if not self.model.iter_is_valid(iter):
            return

        filenum, page_num, rotation = self.model.get(iter, 2, 3, 6)
        crop = self.model.get(iter, 7, 8, 9, 10)
        page = self.pdfqueue[filenum-1].document.get_page(page_num-1)
        w0, h0 = page.get_size()

        rotation = int(rotation) % 360
        rotation = round(rotation / 90) * 90
        if rotation == 90 or rotation == 270:
            w1, h1 = h0, w0
        else:
            w1, h1 = w0, h0

        self.model.set(iter, 11, w1, 12, h1)


    def progress_bar_timeout(self):
        """Function for progress bar."""

        cnt_finished = 0
        cnt_all = 0
        for row in self.model:
            cnt_all += 1
            if row[1]:
                cnt_finished += 1
        fraction = float(cnt_finished)/float(cnt_all)

        self.progress_bar.set_fraction(fraction)
        self.progress_bar.set_text('Rendering thumbnails... [%(i1)s/%(i2)s]'
                                   % {'i1' : cnt_finished, 'i2' : cnt_all})
        if fraction >= 0.999:
            self.progress_bar.hide()
            return False
        elif not self.progress_bar.get_visible():
            self.progress_bar.show()

        return True

    def update_thumbnail(self, object, num, thumbnail):
        """Function to update the thumbnail."""

        row = self.model[num]
        row[1] = thumbnail
        row[4] = self.zoom_scale

    def zoom_set(self, level):
        """Function to set the zoom level."""

        self.zoom_level = max(min(level, 5), -24)
        self.zoom_scale = 1.1 ** self.zoom_level
        for row in self.model:
            row[4] = self.zoom_scale

    def zoom_change(self, step=5):
        """Function to modify the zoom level."""

        self.zoom_set(self.zoom_level + step)


def main():
    GObject.threads_init()
    PDFShuffler()
    Gtk.main()

if __name__ == "__main__":
    main()
