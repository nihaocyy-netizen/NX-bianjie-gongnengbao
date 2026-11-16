# -*- coding: utf-8 -*-
import os
import sys
import importlib
import NXOpen


def main():
    directory = os.path.dirname(os.path.abspath(__file__))
    if directory not in sys.path:
        sys.path.insert(0, directory)
    module = importlib.import_module("nx_screw_insert_ui")
    module.NativeScrewDialog().show()


def GetUnloadOption(dummy):
    return NXOpen.Session.LibraryUnloadOption.AtTermination


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        NXOpen.UI.GetUI().NXMessageBox.Show("Screw Parameters",
            NXOpen.NXMessageBox.DialogType.Error, str(error))
