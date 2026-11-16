import os
import sys
import NXOpen


def main():
    directory=os.path.dirname(os.path.abspath(__file__))
    if directory not in sys.path:
        sys.path.insert(0,directory)
    import nx_screw_native
    nx_screw_native.register()


def GetUnloadOption(dummy):
    return NXOpen.Session.LibraryUnloadOption.AtTermination


if __name__=='__main__':
    main()
