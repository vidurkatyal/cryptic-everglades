try:
    import gi
    gi.require_version('Gtk', '3.0')
    from gi.repository import Gtk
except ValueError, e:
    raise e

import sys
import os
import mimetypes
from gi.repository import Gdk


class PDFShuffler:
    """Class for PDFShuffler."""

    def __init__(self):
        self.uiBuilder = Gtk.Builder()
        self.uiBuilder.add_from_file("GUI.glade")
        self.uiBuilder.connect_signals(self)

        self.window = self.uiBuilder.get_object("main_window")
        self.window.set_title("PDFShuffler")
        self.window.set_border_width(0)
        self.window.set_default_size(
            min(700, Gdk.Screen.get_default().get_width() / 2),
            min(600, Gdk.Screen.get_default().get_height() - 50)
            )
        self.window.connect('delete-event', self.close_window)



        self.window.show_all()

        self.import_directory = self.export_directory = os.getenv('HOME')


    def close_window(self, widget, event=None, data=None):
        """Handler for closing the application."""

        if Gtk.main_level():
            Gtk.main_quit()
        else:
            sys.exit(0)

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
                    errors += "File %s not supported! Hence not importing..\n" % filename
                    continue
                print "File %s qualifies" % filename
                # TODO: Add the file to application

            if errors:
                self.error_message_dialog(errors)

        file_import.destroy()

    def export_pdf(self, widget=None, data=None):
        """Handler for exporting the pdf."""

        # TODO: Checks to see that the application currently has some pdfs

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

            print "Exporting file %s" % _file
            # TODO: Export the file to pdf

        file_export.destroy()




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


def main():
    PDFShuffler()
    Gtk.main()

if __name__ == "__main__":
    main()
