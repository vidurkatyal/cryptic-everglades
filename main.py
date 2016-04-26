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
import urllib
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

        self.sw.connect('drag_data_received', self.sw_dnd_received_data)

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

        self.iconview.connect('button_press_event', self.iconview_button_press_event)
        self.iconview.connect('drag_data_get', self.iconview_dnd_get_data)
        self.iconview.connect('drag_data_received', self.iconview_dnd_received_data)
        self.iconview.connect('drag_data_delete', self.iconview_dnd_data_delete)
        self.iconview.connect('drag_motion', self.iconview_dnd_motion)
        self.iconview.connect('drag_leave', self.iconview_dnd_leave_end)
        self.iconview.connect('drag_end', self.iconview_dnd_leave_end)




        align.add(self.iconview)


        # Popup menu for iconview
        self.popup = Gtk.Menu()
        labels = ('Rotate Right', 'Rotate Left', 'Delete')
        handlers = (self.rotate_page_right, self.rotate_page_left, self.remove_selected_pages)
        for label, handler in zip(labels, handlers):
           popup_item = Gtk.MenuItem.new_with_mnemonic(label)
           popup_item.connect('activate', handler)
           popup_item.show()
           self.popup.append(popup_item)

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
        self.window.connect('key_press_event', self.on_keypress)
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


    def on_keypress(self, widget, event):
        """Handler to detect keypress."""

        if event.keyval == Gdk.KEY_Delete:
            self.remove_selected_pages()

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
        if not len(self.model):
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


    def remove_selected_pages(self, button=None):
        """Handler to remove selected pages from the application."""

        model = self.iconview.get_model()
        selection = self.iconview.get_selected_items()
        selection.sort(reverse=True)
        for path in selection:
            iter = model.get_iter(path)
            model.remove(iter)

    def zoom_in(self, widget=None):
        """Handler to increase the zoom level by 5 steps"""

        self.zoom_change(5)

    def zoom_out(self, widget=None, step=5):
        """Handler to reduce the zoom level by 5 steps."""

        self.zoom_change(-5)

    def rotate_page_right(self, widget, data=None):
        """Handler to rotate selected pages right."""

        self.rotate_page(90)

    def rotate_page_left(self, widget, data=None):
        """Handler to rotate selected pages left."""

        self.rotate_page(-90)


    def rotate_page(self, angle):
        """Function to rotate selected pages."""

        model = self.iconview.get_model()
        selection = self.iconview.get_selected_items()

        for path in selection:
            iter = model.get_iter(path)
            new_angle = ( model.get_value(iter, 6) + int(angle) ) % 360
            model.set_value(iter, 6, new_angle)
            self.update_geometry(iter)
        self.reset_iconview_width()

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
            angle = row[6]
            current_page.rotateClockwise(angle)
            pdf_output.addPage(current_page)

        pdf_output.write(open(_file, 'wb'))

        (path, filename) = os.path.split(_file)
        info = "PDF " + filename + " exported successfully."
        self.info_message_dialog(info)



    def info_message_dialog(self, msg):
        """Function to display info messages."""

        dialog = Gtk.MessageDialog(flags=Gtk.DialogFlags.MODAL,
                                          type=Gtk.MessageType.INFO,
                                          parent=self.window,
                                          message_format=str(msg),
                                          buttons=Gtk.ButtonsType.OK)
        response = dialog.run()
        if response == Gtk.ResponseType.OK:
            dialog.destroy()


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


    def sw_dnd_received_data(self, scrolledwindow, context, x, y,
                             selection_data, target_id, etime):
        """Handler to import files by by drag and drop in scrolledwindow."""

        data = selection_data.get_data()
        if target_id == self.TEXT_URI_LIST:
            uri = data.strip()
            uris = uri.split() # For multiple files dropped
            errors = ""
            for uri in uris:
                _file = self.get_file_path_from_dnd_dropped_uri(uri)
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


    def iconview_button_press_event(self, iconview, event):
        """Handler for managing mouse clicks on the iconview"""

        button = event.button
        if button == 3:
            x = int(event.x)
            y = int(event.y)
            time = event.time
            path = iconview.get_path_at_pos(x, y)
            selection = iconview.get_selected_items()
            if path:
                if path not in selection:
                    iconview.unselect_all()
                iconview.select_path(path)
                iconview.grab_focus()
                self.popup.popup(None, None, None, None, button, time)
            return 1


    def iconview_dnd_get_data(self, iconview, context,
                        selection_data, target_id, etime):
        """Handler to get the pages selected by drag and drop in iconview."""

        model = iconview.get_model()
        selection = self.iconview.get_selected_items()
        selection.sort(key=lambda x: x.get_indices()[0])
        data = []
        for path in selection:
            target = str(selection_data.get_target())
            if target == 'MODEL_ROW_INTERN':
                data.append(str(path[0]))

        if data:
            data = '\n;\n'.join(data)
            selection_data.set(selection_data.get_target(), 8, data.encode())


    def iconview_dnd_received_data(self, iconview, context, x, y,
                             selection_data, target_id, etime):
        """Handler to receive pages sent by drag and drop in iconview."""

        model = iconview.get_model()
        data = selection_data.get_data()
        if data:
            data = data.decode().split('\n;\n')
            item = iconview.get_dest_item_at_pos(x, y)
            if item:
                path, position = item
                ref_to = Gtk.TreeRowReference.new(model,path)
            else:
                ref_to = None
                position = Gtk.IconViewDropPosition.DROP_RIGHT
                if len(model) > 0:  #find the iterator of the last row
                    row = model[-1]
                    ref_to = Gtk.TreeRowReference.new(model, row.path)
            if ref_to:
                before = (position == Gtk.IconViewDropPosition.DROP_LEFT
                          or position == Gtk.IconViewDropPosition.DROP_ABOVE)
                target = str(selection_data.get_target())

                if target == 'MODEL_ROW_INTERN':
                    if before:
                        data.sort(key=int)
                    else:
                        data.sort(key=int,reverse=True)
                    ref_from_list = [Gtk.TreeRowReference.new(model, Gtk.TreePath(p))
                                     for p in data]
                    for ref_from in ref_from_list:
                        path = ref_to.get_path()
                        iter_to = model.get_iter(path)
                        path = ref_from.get_path()
                        iter_from = model.get_iter(path)
                        row = model[iter_from]
                        if before:
                            model.insert_before(iter_to, row[:])
                        else:
                            model.insert_after(iter_to, row[:])
                    if context.get_actions() & Gdk.DragAction.MOVE:
                        for ref_from in ref_from_list:
                            path = ref_from.get_path()
                            iter_from = model.get_iter(path)
                            model.remove(iter_from)


    def iconview_dnd_data_delete(self, widget, context):
        """Handler that deletes the drag and drop items after a successful move operation."""

        model = self.iconview.get_model()
        selection = self.iconview.get_selected_items()
        ref_del_list = [Gtk.TreeRowReference.new(model,path) for path in selection]
        for ref_del in ref_del_list:
            path = ref_del.get_path()
            iter = model.get_iter(path)
            model.remove(iter)


    def iconview_dnd_motion(self, iconview, context, x, y, etime):
        """Handler that initiates auto-scroll during drag and drop."""

        autoscroll_area = 40
        sw_vadj = self.sw.get_vadjustment()
        sw_height = self.sw.get_allocation().height
        if y -sw_vadj.get_value() < autoscroll_area:
            if not self.iconview_auto_scroll_timer:
                self.iconview_auto_scroll_direction = Gtk.DirectionType.UP
                self.iconview_auto_scroll_timer = GObject.timeout_add(150,
                                                                self.iconview_auto_scroll)
        elif y -sw_vadj.get_value() > sw_height - autoscroll_area:
            if not self.iconview_auto_scroll_timer:
                self.iconview_auto_scroll_direction = Gtk.DirectionType.DOWN
                self.iconview_auto_scroll_timer = GObject.timeout_add(150,
                                                                self.iconview_auto_scroll)
        elif self.iconview_auto_scroll_timer:
            GObject.source_remove(self.iconview_auto_scroll_timer)
            self.iconview_auto_scroll_timer = None


    def iconview_dnd_leave_end(self, widget, context, ignored=None):
        """Handler that ends the auto-scroll during drag and drop."""

        if self.iconview_auto_scroll_timer:
            GObject.source_remove(self.iconview_auto_scroll_timer)
            self.iconview_auto_scroll_timer = None


    def iconview_auto_scroll(self):
        """Function for timeout for auto-scroll."""

        sw_vadj = self.sw.get_vadjustment()
        sw_vpos = sw_vadj.get_value()
        if self.iconview_auto_scroll_direction == Gtk.DirectionType.UP:
            sw_vpos -= sw_vadj.get_step_increment()
            sw_vadj.set_value(max(sw_vpos, sw_vadj.get_lower()))
        elif self.iconview_auto_scroll_direction == Gtk.DirectionType.DOWN:
            sw_vpos += sw_vadj.get_step_increment()
            sw_vadj.set_value(min(sw_vpos, sw_vadj.get_upper() - sw_vadj.get_page_size()))
        return True


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
        self.reset_iconview_width()

    def zoom_change(self, step=5):
        """Function to modify the zoom level."""

        self.zoom_set(self.zoom_level + step)

    def get_file_path_from_dnd_dropped_uri(self, uri):
        """Extracts the path from an uri"""

        path = urllib.url2pathname(uri) # escape special chars
        path = path.strip('\r\n\x00')   # remove \r\n and NULL

        if path.startswith('file:\\\\\\'): # windows
            path = path[8:]
        elif path.startswith('file://'):   # nautilus, rox
            path = path[7:]
        elif path.startswith('file:'):     # xffm
            path = path[5:]  # 5 is len('file:')

        return path


def main():
    GObject.threads_init()
    PDFShuffler()
    Gtk.main()

if __name__ == "__main__":
    main()
