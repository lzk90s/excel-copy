import inspect
import os


class defer:
    """
    Proof of concept for a python equivalent of golang's defer statement

    Note that the callback order is probably not guaranteed

    """

    def __init__(self, callback, *args, **kwargs):
        self.callback = callback
        self.args = args
        self.kwargs = kwargs

        # Add a reference to self in the caller variables so our __del__
        # method will be called when the function goes out of scope
        caller = inspect.currentframe().f_back
        caller.f_locals[b'_' + os.urandom(48)] = self

    def __del__(self):
        self.callback(*self.args, **self.kwargs)
