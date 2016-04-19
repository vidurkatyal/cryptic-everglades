try:
    import gi
    gi.require_version('Gtk', '3.0')
    from gi.repository import Gtk
except ValueError, e:
    raise e

import sys
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


    def close_window(self, widget, event=None, data=None):
        if Gtk.main_level():
            Gtk.main_quit()
        else:
            sys.exit(0)


def main():
    PDFShuffler()
    Gtk.main()

if __name__ == "__main__":
    main()
